"""G1b (schema v11): blocks are found by code; the model describes each distinct block shape once (<= 3 patterns).

record_blocks and column_spans are gone from the response schema; merged-range spans come from code. Older plan
shapes (explicit record_blocks) stay accepted by the validator as legacy input and keep their own tests.
"""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path

import pytest

from carthagent.mapper import MAPPER_PROMPT_VERSION, MAPPER_SCHEMA_VERSION, MapperConfig, MapperValidationError, OpenAICompatibleMapper

ROOT = Path(__file__).resolve().parents[2]


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


adv = _load(Path(__file__).with_name("test_option_b_adversarial.py"), "g1b_adversarial")
layouts = _load(Path(__file__).with_name("test_option_b_layouts.py"), "g1b_layouts")
S = "Anlagengitter"


def v11(plan_name, patterns=None, drop=("record_blocks", "column_spans")):
    """The legacy recorded plan rewritten in the v11 shape: no record_blocks/column_spans, patterns instead."""
    recorded = copy.deepcopy(adv.plan(plan_name))
    for q in recorded["plans"]:
        for key in drop:
            q.pop(key, None)
        q["block_pattern"] = patterns or []
    return recorded


def pattern(rows, role="data", row_count=4):
    return {"rows": [{"ordinal": o, "header_row_id": (f"@row:{S}:{h}" if h else "none"), "measure": m or "none"} for o, h, m in rows], "role": role,
            "row_count": row_count, "applies_to": "all_blocks_of_same_shape"}


SAP = pattern([(1, 16, "ahk"), (2, 17, "cum_depreciation"), (3, 18, "book_value")])


def test_g1b_schema_has_no_record_blocks_or_spans_and_no_unions():
    schema = OpenAICompatibleMapper.response_schema()
    plan = schema["properties"]["plans"]["items"]["properties"]
    assert "record_blocks" not in plan and "column_spans" not in plan
    assert schema["$defs"]["block_pattern"]["required"] == ["rows", "role", "row_count", "applies_to"]
    text = json.dumps(schema, separators=(",", ":"))
    assert "anyOf" not in text and '"null"' not in text, "nullables are the sentinel 'none'"
    assert (MAPPER_SCHEMA_VERSION, MAPPER_PROMPT_VERSION) == ("semantic_region_plan_v11", "compact-region-plan-v13")


def test_g1b_sap_pattern_binds_the_fixture(tmp_path):
    source = adv.fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    _, facts = adv.run(tmp_path, source, v11("sap_block_unlabelled", [SAP]))
    assert layouts.facts_as_rows(facts) == layouts.expected_sap()


def test_g1b_spiegel_binds_through_code_derived_merged_spans(tmp_path):
    _, facts = adv.run(tmp_path, adv.fixtures.two_year_two_sheets(tmp_path / "s.xlsx"), v11("two_year_two_sheets"))
    assert layouts.facts_as_rows(facts) == layouts.expected_spiegel()


def test_g1b_flat_register_binds_without_blocks(tmp_path):
    recorded = v11("flat_sap_register")
    for q in recorded["plans"]:
        q["unit_contexts"] = []
    _, facts = adv.run(tmp_path, adv.fixtures.flat_sap_register(tmp_path / "r.xlsx"), recorded)
    assert layouts.facts_as_rows(facts) == layouts.expected_register()


@pytest.mark.parametrize("plan_name", ["flat_sap_register", "two_year_two_sheets"])
def test_g1b_one_row_records_bind_through_a_row_count_one_pattern(tmp_path, plan_name):
    header = {"flat_sap_register": ("Sheet1", 13), "two_year_two_sheets": None}[plan_name]
    recorded = v11(plan_name)
    for q in recorded["plans"]:
        sheet = q["region_observation_id"].split(":", 1)[1]
        row = header[1] if header else 5
        q["block_pattern"] = [{"rows": [{"ordinal": 0, "header_row_id": f"@row:{sheet}:{row}", "measure": "none"}], "role": "data", "row_count": 1, "applies_to": "all_blocks_of_same_shape"}]
    builder = {"flat_sap_register": adv.fixtures.flat_sap_register, "two_year_two_sheets": adv.fixtures.two_year_two_sheets}[plan_name]
    _, facts = adv.run(tmp_path, builder(tmp_path / "x.xlsx"), recorded)
    expected = {"flat_sap_register": layouts.expected_register, "two_year_two_sheets": layouts.expected_spiegel}[plan_name]()
    assert layouts.facts_as_rows(facts) == expected


@pytest.mark.parametrize("name", ["swapped_header_rows", "shifted_header_rows", "measure_swapped_consistently", "null_measures"])
def test_g1b_adversarial_patterns_bind_nothing(tmp_path, name):
    rows = {"swapped_header_rows": [(1, 17, "ahk"), (2, 16, "cum_depreciation"), (3, 18, "book_value")],
            "shifted_header_rows": [(1, 17, "ahk"), (2, 18, "cum_depreciation"), (3, None, "book_value")],
            "measure_swapped_consistently": [(1, 17, "cum_depreciation"), (2, 16, "ahk"), (3, 18, "book_value")],
            "null_measures": [(1, 16, None), (2, 17, None), (3, 18, None)]}[name]
    _, facts = adv.run(tmp_path, adv.fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), v11("sap_block_unlabelled", [pattern(rows)]))
    assert not facts


def test_g1b_numeric_identity_cell_never_keys_a_block(tmp_path):
    recorded = v11("adversarial_numeric_identity", [SAP])
    _, facts = adv.run(tmp_path, adv.fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx", class_cell=1000.5), recorded)
    assert not [f for f in facts if f["pointer"]["cell"][1:] in ("20", "21", "22")] and not adv.wrong(facts, adv.sap_truth())


def test_g1b_a_block_matching_two_patterns_or_none_abstains(tmp_path):
    source = adv.fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    _, facts = adv.run(tmp_path, source, v11("sap_block_unlabelled", [SAP, copy.deepcopy(SAP)]), "two")
    assert not facts
    _, facts = adv.run(tmp_path, source, v11("sap_block_unlabelled", [pattern([(1, 16, "ahk"), (2, 17, "cum_depreciation")], row_count=3)]), "none")
    assert not facts


def test_g1b_more_than_three_patterns_is_malformed():
    from carthagent.acquire_xlsx import acquire_xlsx
    from carthagent.policies import DEFAULT_BOUNDS
    mapper = OpenAICompatibleMapper(MapperConfig(True, "http://127.0.0.1:1/v1/responses", "m", 1, 1)); mapper.last_verification = []
    plan = {"plans": [{"region_observation_id": "d:x", "table_type": "unknown", "header_rows": [], "block_pattern": [SAP] * 4, "field_headers": [], "identity_headers": [],
                       "measure_axes": [], "movement_axes": [], "row_roles": [], "period_contexts": [], "unit_contexts": []}]}
    with pytest.raises(MapperValidationError):
        mapper._validate("d", (), plan)
