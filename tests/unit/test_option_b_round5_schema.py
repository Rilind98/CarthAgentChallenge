"""Fix round 5 unit tests: G1 grammar-sized schema, G2 stacked-header exception confirmed by local code."""
from __future__ import annotations

import json

from carthagent.mapper import MAPPER_SCHEMA_VERSION, OpenAICompatibleMapper
from carthagent.structure import FIELDS, IDENTITIES

FORBIDDEN = {"minItems", "maxItems", "minimum", "maximum", "minLength", "maxLength", "pattern", "format", "default"}


def _nodes(node, defs, seen=()):
    if isinstance(node, dict):
        if "$ref" in node:
            name = node["$ref"].rsplit("/", 1)[1]
            assert name in defs, node["$ref"]
            if name not in seen:
                yield from _nodes(defs[name], defs, (*seen, name))
            return
        yield node
        for key in ("properties",):
            for child in node.get(key, {}).values():
                yield from _nodes(child, defs, seen)
        if "items" in node:
            yield from _nodes(node["items"], defs, seen)
        for child in node.get("anyOf", ()):
            yield from _nodes(child, defs, seen)


def test_g1_schema_states_every_enum_once_and_stays_strict():
    schema = OpenAICompatibleMapper.response_schema()
    text = json.dumps(schema, separators=(",", ":"))
    defs = schema["$defs"]
    enums = [node["enum"] for node in json.loads(text)["$defs"].values() if isinstance(node, dict) and "enum" in node]
    for values in (sorted(FIELDS), sorted(IDENTITIES)):
        assert text.count(json.dumps(values, separators=(",", ":"))) == 1
    assert text.count('"enum"') == len(enums) + 1, "every enum lives once in $defs (plus the one-value applies_to marker)"
    assert text.count("anyOf") == 0, "v11: the nullables are the sentinel 'none', no unions remain"
    assert len(text) < 5400, len(text)
    for node in _nodes(schema, defs):
        assert not FORBIDDEN & set(node)
        if node.get("type") == "object":
            assert node["additionalProperties"] is False and set(node["required"]) == set(node["properties"])
    plan = schema["properties"]["plans"]["items"]
    assert plan["properties"]["block_pattern"]["type"] == "array", "no object|null union for the pattern"
    assert MAPPER_SCHEMA_VERSION == "semantic_region_plan_v11"


def _two_row_header_sheet(tmp_path):
    from openpyxl import Workbook
    source = tmp_path / "two_row.xlsx"; workbook = Workbook(); sheet = workbook.active; sheet.title = "S"
    sheet["A1"] = "Anlagenspiegel Geschäftsjahr 2025 EUR"
    sheet["C3"] = "Anschaffungskosten"; sheet.merge_cells("C3:D3")
    sheet["A4"] = "Anlage"; sheet["C4"] = "01.01.2025"; sheet["D4"] = "31.12.2025"
    sheet["A5"] = "A-1"; sheet["C5"] = 100; sheet["D5"] = 110
    workbook.save(source); return source


def test_g2_plan_cannot_override_a_correct_local_field_by_naming_two_header_rows(tmp_path):
    from carthagent.acquire_xlsx import acquire_xlsx
    from carthagent.models import AssociationClaim, AssociationKind
    from carthagent.policies import DEFAULT_BOUNDS
    from carthagent.semantics import _apply_mapped_axes
    from carthagent.structure import ContextResolver
    observations = acquire_xlsx(_two_row_header_sheet(tmp_path), "d", "a" * 64, DEFAULT_BOUNDS).observations
    by = {o.locator: o for o in observations}
    amount = by["S!C5"]; context = ContextResolver(observations).resolve(amount, observations)
    assert context.business.field == "ahk_open"
    row = lambda n: f"d:xlsx:S:row:{n}"  # noqa: E731
    claim = AssociationClaim(AssociationKind.OTHER, (row(5),), (row(5),), (), True, "ahk_close", (by["S!C4"].id,), (), None, None, None, None, "data",
                             (row(3), row(4)), by["S!C4"].id, row(4), "ahk", None, None, row(5))
    applied = _apply_mapped_axes(context, claim, amount, observations)
    assert "field" in applied.missing_axes and applied.business.field == "ahk_open", "a caption row plus one leaf row is not a stacked header"


def test_g2_sap_stacked_header_still_defers_to_the_block(tmp_path):
    import importlib.util
    from decimal import Decimal
    from pathlib import Path
    root = Path(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("g2_r4", root / "tests" / "unit" / "test_option_b_round4.py"); r4 = importlib.util.module_from_spec(spec); spec.loader.exec_module(r4)
    source, truth = r4.fixtures.sap_block_many(tmp_path / "many.xlsx", blocks=6)
    facts, _ = r4._facts(tmp_path, source, r4.pattern_plan())
    assert r4._rows(facts) == {cell: (field, klass, Decimal(str(value))) for cell, (field, klass, value) in truth.items()}
