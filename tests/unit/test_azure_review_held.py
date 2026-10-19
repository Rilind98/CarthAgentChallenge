"""Code-review fixes that land with the held semantics.py/structure.py hunks (C-B1, C-B2, C-S1).

Probes are built from the recorded synthetic responses (real Azure geometry,
edited the way Azure is known to behave) and replayed through the loopback stub.
"""
from __future__ import annotations

import copy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_azure_layout import (  # noqa: E402
    KEY_ENV, LayoutStub, acquire, cells_of, documents, recorded, stub,  # noqa: F401  (fixtures)
)

from carthagent.policies import DEFAULT_POLICIES  # noqa: E402
from carthagent.semantics import interpret  # noqa: E402


def _facts(result):
    _, facts, abstentions, _ = interpret(result.observations, result.coverage, DEFAULT_POLICIES)
    return facts, abstentions


def _cell(result, table: int, row: int, column: int):
    return next(c for c in cells_of(result) if (c.metadata["table_index"], c.metadata["row_index"], c.metadata["column_index"]) == (table, row, column))


def _with_second_table(payload: dict, *, header_text: dict[int, str], drop_header: bool, dy: float = 3.0) -> dict:
    """Append a copy of table 0 (and its words) shifted down by dy inches as table 1."""
    result = payload["analyzeResult"]
    page = result["pages"][0]
    content = result["content"]
    moved_offsets: dict[int, int] = {}
    new_words = []
    for word in page["words"]:
        polygon = [value + (dy if index % 2 else 0) for index, value in enumerate(word["polygon"])]
        offset = len(content) + 1
        content += " " + word["content"]
        moved_offsets[word["span"]["offset"]] = offset
        new_words.append({**copy.deepcopy(word), "polygon": polygon, "span": {"offset": offset, "length": len(word["content"])}})
    table = copy.deepcopy(result["tables"][0])
    cells = []
    for cell in table["cells"]:
        if drop_header and cell["rowIndex"] == 0:
            continue
        for region in cell["boundingRegions"]:
            region["polygon"] = [value + (dy if index % 2 else 0) for index, value in enumerate(region["polygon"])]
        cell["spans"] = [{"offset": moved_offsets.get(span["offset"], span["offset"]), "length": span["length"]} for span in cell.get("spans", [])]
        if cell["rowIndex"] == 0 and cell["columnIndex"] in header_text:
            text = header_text[cell["columnIndex"]]
            word = next(w for w in new_words if w["span"]["offset"] == cell["spans"][0]["offset"])
            word["content"] = cell["content"] = text
        cells.append(cell)
    table["cells"] = cells
    for region in table["boundingRegions"]:
        region["polygon"] = [value + (dy if index % 2 else 0) for index, value in enumerate(region["polygon"])]
    page["words"] = page["words"] + new_words
    result["content"] = content
    result["tables"] = result["tables"] + [table]
    return payload


def test_lower_table_with_its_own_header_never_uses_the_upper_header(stub, documents, monkeypatch):
    # C-B1: the lower table's column 3 is headed "Abschreibungen", which has no
    # field alias; the richer header of the upper table must not lend it "Zugänge".
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(_with_second_table(recorded("scanned_p1"), header_text={1: "Text", 3: "Abschreibungen"}, drop_header=False))
    result = acquire(documents / "azure_scanned.pdf", stub)
    lower = _cell(result, 1, 1, 3)
    assert lower.text == "120,00" and lower.metadata["verified"] is True
    facts, abstentions = _facts(result)
    assert not [f for f in facts if f.source_observation_id == lower.id]
    assert any(lower.id in a.candidate_observation_ids and "column_structure" in a.reason for a in abstentions)
    lower_fields = {f.business.field for f in facts if ":azcell:1:" in f.source_observation_id}
    assert "zugang" not in lower_fields and lower_fields, lower_fields
    upper = {f.business.field for f in facts if ":azcell:0:" in f.source_observation_id}
    assert "zugang" in upper


def test_table_without_its_own_header_row_abstains(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(_with_second_table(recorded("scanned_p1"), header_text={}, drop_header=True))
    result = acquire(documents / "azure_scanned.pdf", stub)
    lower_cells = [c for c in cells_of(result) if c.metadata["table_index"] == 1 and c.text and c.text[0].isdigit() and "," in c.text]
    assert len(lower_cells) == 10 and all(c.metadata["verified"] for c in lower_cells)
    facts, _ = _facts(result)
    assert not [f for f in facts if ":azcell:1:" in f.source_observation_id]
    assert len([f for f in facts if ":azcell:0:" in f.source_observation_id]) == 10


def test_printed_digits_not_the_cell_polygon_choose_the_column(stub, documents, monkeypatch):
    # C-B2: Azure widened the 1.000,00 cell so its polygon centre lies under
    # "Zugänge" while the digits stay under "AHK 01.01.2025".
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    cell = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["rowIndex"] == 1 and c["columnIndex"] == 2)
    cell["boundingRegions"][0]["polygon"] = [4.3, 1.2, 5.85, 1.19, 5.85, 1.48, 4.3, 1.48]
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    widened = _cell(result, 0, 1, 2)
    assert widened.metadata["verified"] is True
    polygon_centre = (widened.bbox[0] + widened.bbox[2]) / 2
    assert polygon_centre > 5.0 * 72, "the probe must put the polygon centre under the next header"
    facts, _ = _facts(result)
    assert [f.business.field for f in facts if f.source_observation_id == widened.id] == ["ahk_open"]
    assert sorted(f.business.field for f in facts if f.business.asset_or_class == "000101") == ["abgang", "ahk_close", "ahk_open", "bw_close", "zugang"]


def test_numeric_cell_spanning_two_columns_abstains(stub, documents, monkeypatch):
    # C-B2: on the rotated scan Azure gave "2.000,00" columnSpan 2; label the
    # empty neighbour header "Zugänge" as the review probe did.
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_rotated_90_p1")
    header = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["rowIndex"] == 0 and c["columnIndex"] == 3)
    header["content"] = "Zugänge"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned_rotated_90.pdf", stub)
    spanning = _cell(result, 0, 2, 2)
    assert spanning.text == "2.000,00" and spanning.metadata["column_span"] == 2 and spanning.metadata["verified"] is True
    facts, abstentions = _facts(result)
    assert not [f for f in facts if f.source_observation_id == spanning.id]
    assert [a.reason for a in abstentions if spanning.id in a.candidate_observation_ids] == ["ambiguous_ocr_cell:column_span"]
    assert not any(f.business.field == "zugang" and f.business.asset_or_class == "000202" and str(f.amount.source_signed_value) == "2000.00" for f in facts)


def test_words_of_a_cell_that_failed_the_cross_check_abstain_too(stub, documents, monkeypatch):
    # C-S1: the cell text disagrees with its words; neither the cell nor its
    # word may become a fact (the word path's legacy layout would otherwise run).
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    cell = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["content"] == "1.000,00")
    cell["content"] = "7.000,00"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    bad = next(c for c in cells_of(result) if c.text == "7.000,00")
    assert bad.metadata["verified"] is False
    word_id = bad.metadata["member_word_ids"][0]
    facts, abstentions = _facts(result)
    assert not [f for f in facts if f.source_observation_id in {bad.id, word_id}]
    assert [a.reason for a in abstentions if word_id in a.candidate_observation_ids] == ["unverified_ocr_cell:text_mismatch"]
    assert len(facts) == 9
