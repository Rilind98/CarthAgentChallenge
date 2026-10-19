"""Option C part B: how semantics/structure consume Azure cells.

Lands with the part B hunks of semantics.py and structure.py. Replays recorded
Azure responses through the loopback stub of test_azure_layout.py.
"""
from __future__ import annotations

from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_azure_layout import (  # noqa: E402
    KEY_ENV, LayoutStub, acquire, cells_of, contains, documents, recorded, stub,  # noqa: F401  (fixtures)
)

from carthagent.policies import DEFAULT_POLICIES  # noqa: E402
from carthagent.semantics import _pointer_for, interpret  # noqa: E402
from carthagent.structure import ContextResolver  # noqa: E402


def _cell(result, row: int, column: int):
    return next(c for c in cells_of(result) if c.metadata["row_index"] == row and c.metadata["column_index"] == column)


def test_verified_cell_pointer_is_the_lexeme_and_keeps_the_cell_polygon(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_scanned.pdf", stub)
    cell = _cell(result, 1, 2)
    assert cell.text == "1.000,00" and cell.metadata["verified"] is True
    pointer = _pointer_for(cell)
    assert pointer["bbox"] == list(cell.metadata["lexeme_bbox"])
    assert pointer["cell_polygon"] == list(cell.metadata["cell_polygon"]) == list(cell.bbox)
    assert pointer["deskew_degrees"] == cell.metadata["deskew_degrees"]
    assert contains(pointer["cell_polygon"], pointer["bbox"], 0.01)
    _, facts, _, _ = interpret(result.observations, result.coverage, DEFAULT_POLICIES)
    assert len(facts) == 10 and all(f.source_observation_id.split(":")[3] == "azcell" for f in facts), "verified cells replace their own words"


def test_row_spanning_numeric_cell_abstains(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    spanning = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["rowIndex"] == 1 and c["columnIndex"] == 2)
    assert spanning["content"] == "1.000,00"
    spanning["rowSpan"] = 2
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    cell = _cell(result, 1, 2)
    assert cell.metadata["verified"] is True and cell.metadata["row_span"] == 2
    _, facts, abstentions, _ = interpret(result.observations, result.coverage, DEFAULT_POLICIES)
    assert not [f for f in facts if f.source_observation_id == cell.id]
    assert [a.reason for a in abstentions if cell.id in a.candidate_observation_ids] == ["ambiguous_ocr_cell:row_span"]
    assert len(facts) == 9


def test_unlabelled_column_abstains_instead_of_inheriting_its_neighbour(stub, documents, monkeypatch):
    # f01 shape: "Abschreibungen (kumuliert)" has no field alias; with word-style
    # axes its values took the field of the labelled column to the left.
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    raw = payload["analyzeResult"]
    header = next(c for c in raw["tables"][0]["cells"] if c["rowIndex"] == 0 and c["columnIndex"] == 3)
    assert header["content"] == "Zugänge"
    word = next(w for w in raw["pages"][0]["words"] if w["content"] == "Zugänge")
    header["content"] = word["content"] = "Abschreibungen"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    resolver = ContextResolver(result.observations)
    unlabelled = _cell(result, 1, 3)
    assert unlabelled.text == "120,00" and unlabelled.metadata["verified"] is True
    context = resolver.resolve(unlabelled, result.observations)
    assert "column_structure" in context.missing_axes
    assert context.business is None or context.business.field in (None, "amount")
    # The labelled neighbour still binds, to its own column only.
    labelled = resolver.resolve(_cell(result, 1, 2), result.observations)
    assert labelled.business.field == "ahk_open" and not labelled.missing_axes
    _, facts, abstentions, _ = interpret(result.observations, result.coverage, DEFAULT_POLICIES)
    assert {f.business.field for f in facts} == {"ahk_open", "abgang", "ahk_close", "bw_close"}
    assert len(facts) == 8
    assert any(unlabelled.id in a.candidate_observation_ids and "column_structure" in a.reason for a in abstentions)


def test_cell_rows_sharing_a_border_still_elect_the_header_above(stub, documents, monkeypatch):
    # On the rotated scan the header row's box touches row 1; electing by row
    # centre keeps the grid header instead of falling back to inherited spans.
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_rotated_90_p1"))
    result = acquire(documents / "azure_scanned_rotated_90.pdf", stub)
    resolver = ContextResolver(result.observations)
    first_row_amount = _cell(result, 1, 7)
    assert first_row_amount.text == "1.100,00"
    context = resolver.resolve(first_row_amount, result.observations)
    assert context.business.field == "ahk_close" and not context.missing_axes
    # Azure split this grid: the value columns 3 and 6 have empty header cells.
    for row, column, text in ((1, 3, "1.000,00"), (1, 6, "20,00"), (2, 6, "40,00")):
        cell = _cell(result, row, column)
        assert cell.text == text
        assert "column_structure" in resolver.resolve(cell, result.observations).missing_axes
