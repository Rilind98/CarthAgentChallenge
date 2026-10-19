"""R3-6: option B on layout-engine (Azure) table cells: azrow logical rows, grid column indices, recorded response.

The scanned f05_03 page 2 grid prints the three SAP header lines merged into ONE grid row (row 1) and the identity
header line as row 0; record blocks are identity row + AHK/AfA/book-value rows. The plan below is hand-written in
the shape the live mapper returns, citing the replayed observations' IDs.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

from carthagent.mapper import MapperConfig, OpenAICompatibleMapper
from carthagent.policies import DEFAULT_POLICIES
from carthagent.semantics import interpret

ROOT = Path(__file__).resolve().parents[2]
_spec = importlib.util.spec_from_file_location("azure_layout_harness", ROOT / "tests" / "unit" / "test_azure_layout.py")
az = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(az)
stub = az.stub  # noqa: F401  (pytest fixture re-export)
CELL = "doc:pdf:p2:azcell:0:{}:{}"
ROW = "doc:pdf:p2:azrow:0:{}"
IDENTITY_ROWS = (2, 14, 22, 34)


def grid_plan(measures=("ahk", "cum_depreciation", "book_value")):
    return {"plans": [{
        "region_observation_id": "doc:pdf:p2:aztable:0", "table_type": "class_grid_matrix",
        "header_rows": [ROW.format(0), ROW.format(1)],
        "record_blocks": [{"identity_row_id": ROW.format(r), "rows": [{"row_id": ROW.format(r + k), "header_row_id": ROW.format(1), "measure": m} for k, m in enumerate(measures, start=1)], "role": "data"}
                          for r in IDENTITY_ROWS],
        "column_spans": [], "field_headers": [],
        "identity_headers": [{"observation_id": CELL.format(0, 6), "identity_role": "class_code"}, {"observation_id": CELL.format(0, 4), "identity_role": "account"}],
        "measure_axes": [],
        "movement_axes": [{"observation_id": CELL.format(1, column), "movement_axis": axis, "parent_header_ids": []}
                          for column, axis in ((0, "opening"), (3, "addition"), (5, "disposal"), (8, "transfer"), (12, "closing"))],
        "row_roles": [], "period_contexts": [],
        "unit_contexts": [{"observation_id": CELL.format(r, 12), "unit": "EUR", "scope": "block"} for r in IDENTITY_ROWS],
    }]}


def _replay(stub_endpoint, tmp_path, monkeypatch, plan):
    monkeypatch.setenv(az.KEY_ENV, "k")
    az.LayoutStub.reset(az.recorded("f05_sap_class_grid_03_p2"))
    observations = az.acquire(az.scan_doc(tmp_path, 2), stub_endpoint).observations
    mapper = OpenAICompatibleMapper(MapperConfig(True, "http://127.0.0.1:1/v1/responses", "replay", 1, 1)); mapper.last_verification = []
    mapped = mapper._validate("doc", observations, plan)
    return observations, mapped, mapper


def test_block_path_reads_grid_rows_and_columns_on_layout_engine_cells(stub, tmp_path, monkeypatch):
    """Since the confidence gate (MIN_NUMERIC_WORD_CONFIDENCE) every record block of this scan holds at least one
    unverified numeric cell, so NO block binds. What must hold: the grid structure is read (identity from the Klasse
    grid column, field from the opening movement x ahk measure), the unverified cells never enter a claim, and each
    affected block abstains on its own (R4 identity / R6 uncheckable) without the plan being voted down."""
    observations, mapped, mapper = _replay(stub, tmp_path, monkeypatch, grid_plan())
    by_id = {o.id: o for o in observations}
    rules = [(item["rule"], item["detail"]) for item in mapper.last_verification]
    assert ("R4", f"identity_incomplete:{ROW.format(22)}") in rules, rules
    assert {detail for rule, detail in rules if rule == "R6"} == {ROW.format(2), ROW.format(14), ROW.format(34)}, rules
    assert not [detail for rule, detail in rules if rule == "R6" and "," in detail], "no plan-level veto: abstention is per block"
    assert not [h for h in mapped if h.candidates[0].record_identity_row_id]
    assert all(by_id[h.amount_observation_id].metadata.get("verified") is True for h in mapped if by_id[h.amount_observation_id].channel == "pdf_ocr" and by_id[h.amount_observation_id].kind == "cell"
               and h.candidates[0].record_identity_row_id)
    assert not interpret(observations, (), DEFAULT_POLICIES, mapped)[1]


def test_unverified_cell_makes_only_its_own_block_abstain(stub, tmp_path, monkeypatch):
    """Mark ONE cell of an otherwise verified synthetic copy unverified: that block abstains, its sibling still verifies."""
    from dataclasses import replace
    observations, _, _ = _replay(stub, tmp_path, monkeypatch, grid_plan())
    # Make every cell verified, then re-break exactly one AfA cell of block 2: the arithmetic of the recorded values decides
    # the rest, so only the per-block outcome for block 2 is asserted.
    patched = tuple(replace(o, metadata={**o.metadata, "verified": True}) if o.kind == "cell" and o.channel == "pdf_ocr" else o for o in observations)
    broken = CELL.format(4, 0)
    patched = tuple(replace(o, metadata={**o.metadata, "verified": False}) if o.id == broken else o for o in patched)
    mapper = OpenAICompatibleMapper(MapperConfig(True, "http://127.0.0.1:1/v1/responses", "replay", 1, 1)); mapper.last_verification = []
    mapped = mapper._validate("doc", patched, grid_plan())
    rules = [(item["rule"], item["detail"]) for item in mapper.last_verification]
    assert ("R6", ROW.format(2)) in rules and not [d for r, d in rules if r == "R6" and "," in d], rules
    assert not [h for h in mapped if h.candidates[0].record_identity_row_id == ROW.format(2)]
    assert broken not in {h.amount_observation_id for h in mapped}


def test_measure_swap_on_the_grid_binds_nothing(stub, tmp_path, monkeypatch):
    observations, mapped, mapper = _replay(stub, tmp_path, monkeypatch, grid_plan(("cum_depreciation", "ahk", "book_value")))
    assert not [h for h in mapped if h.candidates[0].record_identity_row_id]
    assert {"R6"} & {item["rule"] for item in mapper.last_verification}
