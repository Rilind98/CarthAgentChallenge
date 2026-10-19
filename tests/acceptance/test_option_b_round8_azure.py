"""Round 8 R8-5: the PDF straddle rule drops only a true straddle, never a legitimate neighbour ("AfA Abgänge" next to
"Abgänge", "Restbuchwert" next to "Buchwert"). Ported from QA probe test_round7_pdfhdr.py."""

from __future__ import annotations

import collections
import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


azure = _load(ROOT / "tests" / "unit" / "test_azure_layout.py", "r8_azure_layout")
stub = azure.stub  # noqa: F401
documents = azure.documents  # noqa: F401


def relabel(payload, mapping):
    result = payload["analyzeResult"]
    for cell in result["tables"][0]["cells"]:
        if cell["rowIndex"] == 0 and cell["columnIndex"] in mapping:
            text = mapping[cell["columnIndex"]]
            words = [w for w in result["pages"][0]["words"] if any(s["offset"] <= w["span"]["offset"] < s["offset"] + s["length"] for s in cell["spans"])]
            parts = text.split(" ")
            for index, word in enumerate(words):
                word["content"] = " ".join(parts[index:]) if index == len(words) - 1 else (parts[index] if index < len(parts) else "")
            cell["content"] = text
    return payload


CASES = {
    "AfA Abgänge next to Abgänge": ({3: "Abgänge", 4: "AfA Abgänge"}, {3: "abgang", 4: "afa_abgang"}),
    # Undated "Buchwert"/"Restbuchwert" name no field on their own (no date to rank them, same at 32ac60e before the drop
    # rule existed), so neither binds; the point is that neither is DROPPED as a straddle (see the unit test below).
    "Restbuchwert next to Buchwert": ({5: "Buchwert", 6: "Restbuchwert"}, {5: None, 6: None}),
    "true straddle Abgänge Zugänge": ({4: "Abgänge Zugänge"}, {3: "zugang", 4: None}),
}


@pytest.mark.parametrize("name", list(CASES))
def test_r8_5_pdf_header_drop_only_on_a_true_straddle(stub, documents, monkeypatch, name):
    from carthagent.policies import DEFAULT_POLICIES
    from carthagent.semantics import interpret
    monkeypatch.setenv(azure.KEY_ENV, "k"); mapping, expected = CASES[name]
    azure.LayoutStub.reset(relabel(azure.recorded("scanned_p1"), mapping))
    result = azure.acquire(documents / "azure_scanned.pdf", stub)
    facts = interpret(result.observations, result.coverage, DEFAULT_POLICIES)[1]
    by_column = collections.defaultdict(set)
    for fact in facts:
        by_column[int(fact.source_observation_id.rsplit(":", 1)[1])].add(fact.business.field)
    for column, field in expected.items():
        assert by_column.get(column, set()) == ({field} if field else set()), (column, dict(by_column))


@pytest.mark.parametrize(("own", "neighbour", "shared_words", "dropped"), [
    ("AfA Abgänge", "Abgänge", False, False), ("Restbuchwert", "Buchwert", False, False), ("Buchwert Restbuchwert", "Restbuchwert", False, False),
    ("Abgänge Zugänge", "Zugänge", False, True), ("Abschreibungen Abgänge", "Abgänge", True, True),
])
def test_r8_5_straddle_decision(own, neighbour, shared_words, dropped):
    from carthagent.structure import _PdfAxis, _straddles
    axis = _PdfAxis(0, 1, None, None, None, None, own, ("w1", "w2") if shared_words else ("w1",))
    other = _PdfAxis(1, 2, None, None, None, None, neighbour, ("w2",) if shared_words else ("w3",))
    assert _straddles(axis, other) is dropped
