from __future__ import annotations

from decimal import Decimal
import importlib.util
from pathlib import Path

from carthagent.acquire_xlsx import acquire_xlsx
from carthagent.mapper import MAPPER_PROMPT_VERSION, MAPPER_SCHEMA_VERSION, MapperConfig, OpenAICompatibleMapper
from carthagent.pipeline import ACQUISITION_VERSION
from carthagent.policies import DEFAULT_BOUNDS

ROOT = Path(__file__).resolve().parents[2]
_spec = importlib.util.spec_from_file_location("option_b_fixtures", ROOT / "tests" / "fixtures" / "build_option_b.py")
fixtures = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(fixtures)
_stub_spec = importlib.util.spec_from_file_location("option_b_stub", ROOT / "tests" / "fixtures" / "option_b_stub.py")
stub = importlib.util.module_from_spec(_stub_spec); _stub_spec.loader.exec_module(stub)


def test_xlsx_identity_row_followed_by_numeric_rows_is_one_record_block(tmp_path):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    observations = acquire_xlsx(source, "d", "a" * 64, DEFAULT_BOUNDS).observations
    blocks = [o for o in observations if o.kind == "record_block"]
    row = lambda number: f"d:xlsx:Anlagengitter:row:{number}"  # noqa: E731
    assert [list(b.metadata["member_ids"]) for b in blocks] == [[row(19), row(20), row(21), row(22)], [row(24), row(25), row(26), row(27)], [row(29), row(30), row(31), row(32)]]
    assert [b.metadata["identity_row_id"] for b in blocks] == [row(19), row(24), row(29)]
    assert all(b.metadata["sheet"] == "Anlagengitter" and b.channel == "xlsx_structure" for b in blocks)
    assert [b.metadata["subtotal_marker"] for b in blocks] == [False, False, True]
    region = next(o for o in observations if o.kind == "table_region")
    # Blocks are new observations; the region payload stays byte-identical so existing immutable ledgers never collide.
    assert "record_block_ids" not in region.metadata
    assert ACQUISITION_VERSION == "14"


def test_xlsx_flat_register_rows_and_header_rows_do_not_form_record_blocks(tmp_path):
    for builder in (fixtures.flat_sap_register, fixtures.two_year_two_sheets):
        observations = acquire_xlsx(builder(tmp_path / f"{builder.__name__}.xlsx"), "d", "a" * 64, DEFAULT_BOUNDS).observations
        assert not [o.locator for o in observations if o.kind == "record_block"], builder.__name__
        assert all("record_block_ids" not in o.metadata for o in observations if o.kind == "table_region")


FORBIDDEN_SCHEMA_KEYS = {"minItems", "maxItems", "minimum", "maximum", "minLength", "maxLength", "pattern", "format", "default"}


def _walk(schema, path="$"):
    yield path, schema
    for name, child in schema.get("properties", {}).items():
        yield from _walk(child, f"{path}.{name}")
    if "items" in schema:
        yield from _walk(schema["items"], f"{path}[]")
    for index, child in enumerate(schema.get("anyOf", ())):
        yield from _walk(child, f"{path}|{index}")


def test_structure_plan_schema_is_strict_for_openai_and_constraint_free_for_anthropic():
    schema = OpenAICompatibleMapper.response_schema()
    plan = schema["properties"]["plans"]["items"]
    assert {"table_type", "header_rows", "record_blocks", "column_spans"} <= set(plan["required"])
    block_row = plan["properties"]["record_blocks"]["items"]["properties"]["rows"]["items"]
    assert set(block_row["required"]) == {"row_id", "header_row_id", "measure"}
    assert set(plan["properties"]["column_spans"]["items"]["required"]) == {"header_observation_id", "first_column", "last_column"}
    for context in ("period_contexts", "unit_contexts"):
        assert plan["properties"][context]["items"]["properties"]["scope"]["enum"] == ["sheet", "column", "block"]
    for path, node in _walk(schema):
        assert not FORBIDDEN_SCHEMA_KEYS & set(node), path
        if node.get("type") == "object":
            assert node["additionalProperties"] is False and set(node["required"]) == set(node["properties"]), path
    assert (MAPPER_SCHEMA_VERSION, MAPPER_PROMPT_VERSION) == ("semantic_region_plan_v8", "compact-region-plan-v9")


def _mapped(source, plan):
    observations = acquire_xlsx(source, "d", "a" * 64, DEFAULT_BOUNDS).observations
    with stub.recorded_plan_server(stub.load_plan(plan) if isinstance(plan, str) else plan) as endpoint:
        mapper = OpenAICompatibleMapper(MapperConfig(True, endpoint, "local-stub", 2, 256))
        mapped = mapper.map("d", observations)
    return observations, {h.amount_observation_id: h.candidates[0] for h in mapped}, mapper


cell = lambda ref, sheet="Anlagengitter": f"d:xlsx:{sheet}!{ref}"  # noqa: E731
row = lambda number, sheet="Anlagengitter": f"d:xlsx:{sheet}:row:{number}"  # noqa: E731


def test_block_plan_binds_each_measure_row_to_its_declared_header_row(tmp_path):
    _, claims, _ = _mapped(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), "sap_block_unlabelled")
    for ref, field, header_row, measure in (("B20", "ahk_open", 16, "ahk"), ("B21", "kum_afa_open", 17, "cum_depreciation"), ("H22", "bw_close", 18, "book_value"), ("C21", "total_afa_gj", 17, "cum_depreciation")):
        claim = claims[cell(ref)]
        assert (claim.semantic_field, claim.semantic_measure_observation_id, claim.semantic_measure) == (field, row(header_row), measure), ref
        assert claim.record_identity_row_id == row(19) and (cell("E19"), "class_code") in claim.semantic_identity_assignments, ref
        assert claim.target_observation_ids == (row(int(ref[1:])),) and claim.semantic_row_role == "data"
    assert (cell("E24"), "class_code") in claims[cell("B25")].semantic_identity_assignments
    assert cell("B19") not in claims and cell("F20") not in claims, "identity-row cells and header-less columns get no claim"
    assert not any(key.endswith(("B30", "B31", "H32")) for key in claims), "subtotal blocks are declared but never expanded"


# --- T5: verification rules. Every failure must abstain (no claim / no fact), never guess.

def _variant(mutate, name="sap_block_unlabelled"):
    plan = stub.load_plan(name); mutate(plan["plans"][0]); return plan


def _rules(mapper):
    return {item["rule"] for item in mapper.last_verification}


def _blocked(mapped):
    return [h for h in mapped if h.candidates[0].record_identity_row_id is not None]


def _run(source, plan):
    from carthagent.policies import DEFAULT_POLICIES
    from carthagent.semantics import interpret
    observations = acquire_xlsx(source, "d", "a" * 64, DEFAULT_BOUNDS).observations
    with stub.recorded_plan_server(stub.load_plan(plan) if isinstance(plan, str) else plan) as endpoint:
        mapper = OpenAICompatibleMapper(MapperConfig(True, endpoint, "local-stub", 2, 256))
        mapped = mapper.map("d", observations)
    facts = interpret(observations, (), DEFAULT_POLICIES, mapped)[1]
    return mapped, facts, _rules(mapper)


def test_verified_block_facts_take_identity_from_the_identity_row(tmp_path):
    _, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), "sap_block_unlabelled")
    by_cell = {f.physical.occurrence.rsplit("!", 1)[1]: f for f in facts}
    assert not rules
    for ref, field, klass in (("B20", "ahk_open", "1000"), ("B21", "kum_afa_open", "1000"), ("H22", "bw_close", "1000"), ("C26", "total_afa_gj", "2000")):
        business = by_cell[ref].business
        assert (business.field, business.class_code, business.asset_or_class, business.period, business.unit) == (field, klass, klass, "2025", "EUR"), ref
        assert by_cell[ref].source_pointer["cell"] == ref and by_cell[ref].source_pointer["sheet"] == "Anlagengitter"
    assert by_cell["B21"].amount.normalized_value == Decimal("-80000") and by_cell["H22"].amount.normalized_value == Decimal("15000.25")
    assert len(facts) == 26 and not set(by_cell) & {"B19", "F20", "G20", "B30", "H32"}


def test_r1_header_row_must_sit_above_and_blocks_must_be_disjoint(tmp_path):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    def below(plan):
        plan["record_blocks"][0]["rows"][0]["header_row_id"] = "@row:Anlagengitter:22"; plan["header_rows"].append("@row:Anlagengitter:22")
    def overlap(plan):
        plan["record_blocks"][1]["identity_row_id"] = "@row:Anlagengitter:22"
    for mutate in (below, overlap):
        mapped, facts, rules = _run(source, _variant(mutate))
        assert "R1" in rules and not facts, mutate.__name__


def test_r2_non_periodic_blocks_reject_the_whole_plan(tmp_path):
    def ragged(plan):
        plan["record_blocks"][1]["rows"] = plan["record_blocks"][1]["rows"][:2]
    mapped, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), _variant(ragged))
    assert "R2" in rules and not _blocked(mapped) and not facts


def test_r3_field_measure_admissibility_and_one_field_per_header_cell(tmp_path):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    mapped, facts, rules = _run(source, "adversarial_swapped_header_rows")
    assert "R3" in rules and not _blocked(mapped) and not facts
    def relabel(plan):
        plan["field_headers"].append({"observation_id": "@cell:Anlagengitter!C16", "semantic_field": "abgang", "parent_header_ids": []})
    mapped, facts, rules = _run(source, _variant(relabel))
    assert "R3" in rules and not facts


def test_r4_identity_row_must_carry_text_and_identifier_shaped_values(tmp_path):
    mapped, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), "adversarial_shifted_block")
    assert "R4" in rules and not _blocked(mapped) and not facts
    numeric = fixtures.sap_block_unlabelled(tmp_path / "numeric.xlsx", class_cell=1000.5)
    mapped, facts, rules = _run(numeric, "adversarial_numeric_identity")
    first_block = [h for h in _blocked(mapped) if h.candidates[0].record_identity_row_id == row(19)]
    assert "R4" in rules and first_block and all(not h.candidates[0].semantic_identity_assignments for h in first_block)
    assert not [f for f in facts if f.physical.occurrence.endswith(("20", "21", "22"))], "a bare amount in the identity row never becomes the class"
    assert {f.business.class_code for f in facts} == {"2000"}


def test_r5_disagreeing_or_multi_year_period_scopes_abstain(tmp_path):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    def creation_date_too(plan):
        plan["period_contexts"].append({"observation_id": "@cell:Anlagengitter!A13", "temporal_role": "current_period", "scope": "sheet"})
    def two_years(plan):
        plan["period_contexts"] = [{"observation_id": "@cell:Anlagengitter!A13", "temporal_role": "current_period", "scope": "sheet"}]
    for mutate in (creation_date_too,):
        mapped, facts, rules = _run(source, _variant(mutate))
        assert "R5" in rules and _blocked(mapped) and not facts and all(h.candidates[0].semantic_period_observation_id is None for h in _blocked(mapped))
    mapped, facts, rules = _run(source, _variant(two_years))
    assert mapped and len(facts) == 26 and {f.business.period for f in facts} == {"2026"}, "a single year is code-verified text, not a guess"


def test_r6_arithmetic_veto_rejects_structure_that_breaks_most_roll_forwards(tmp_path):
    def swap_open_close(plan):
        labels = {"B16": "ahk_close", "H16": "ahk_open", "B17": "kum_afa_close", "H17": "kum_afa_open"}
        for item in plan["field_headers"]:
            ref = item["observation_id"].rsplit("!", 1)[1]
            item["semantic_field"] = labels.get(ref, item["semantic_field"])
    mapped, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), _variant(swap_open_close))
    assert "R6" in rules and not _blocked(mapped) and not facts


def test_r7_conflicting_row_and_column_measures_abstain_and_unmerged_spans_are_ignored(tmp_path):
    source = fixtures.two_year_two_sheets(tmp_path / "spiegel.xlsx")
    def row_measure(plan):
        for block in plan["record_blocks"]:
            block["rows"][0]["measure"] = "cum_depreciation"
    plan = stub.load_plan("two_year_two_sheets"); row_measure(plan["plans"][0]); row_measure(plan["plans"][1])
    mapped, facts, rules = _run(source, plan)
    cells = {f.physical.occurrence for f in facts}
    assert "R7" in rules and "Tabelle1!C10" not in cells and "Tabelle1!J10" in cells
    def widen(plan):
        plan["column_spans"].append({"header_observation_id": "@cell:Anlagengitter!B16", "first_column": 2, "last_column": 8})
    mapped, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), _variant(widen))
    assert "column_span" in rules and len(facts) == 26


def test_r4_block_marked_subtotal_in_source_cannot_be_declared_data(tmp_path):
    def promote(plan):
        plan["record_blocks"][2]["role"] = "data"
    mapped, facts, rules = _run(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), _variant(promote))
    assert "R4" in rules and len(facts) == 26 and not [f for f in facts if f.physical.occurrence.endswith(("30", "31", "32"))]
