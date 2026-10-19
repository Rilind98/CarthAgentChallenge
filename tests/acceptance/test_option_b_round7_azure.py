"""Round 7 R7-6 (and R7-4 on the layout-engine path): a wide table split across two scanned pages.

Page 1 prints the AHK columns of each asset, page 2 continues the same assets with AfA and book-value columns."""

from __future__ import annotations

import collections
import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


r6 = _load(Path(__file__).with_name("test_option_b_round6_azure.py"), "r7_r6_azure")
azure = r6.azure
stub = azure.stub  # noqa: F401
PAGE_TWO_HEADERS = {2: "AfA 01.01.2025", 3: "AfA Zugänge", 4: "AfA Abgänge", 5: "AfA 31.12.2025", 6: "Buchwert 31.12.2024"}


def _split_width(payload):
    result = payload["analyzeResult"]; table = result["tables"][1]
    for cell in table["cells"]:
        if cell["rowIndex"] == 0 and cell["columnIndex"] in PAGE_TWO_HEADERS:
            text = PAGE_TWO_HEADERS[cell["columnIndex"]]
            words = [w for w in result["pages"][1]["words"] if any(s["offset"] <= w["span"]["offset"] < s["offset"] + s["length"] for s in cell["spans"])]
            for index, word in enumerate(words):
                word["content"] = text if index == 0 else ""
            cell["content"] = text
    return payload


def _run(stub, tmp_path, monkeypatch):
    from carthagent.exporter import TemplateExportAdapter
    from carthagent.policies import DEFAULT_POLICIES
    from carthagent.semantics import interpret
    monkeypatch.setenv(azure.KEY_ENV, "k"); azure.LayoutStub.reset(_split_width(r6._two_pages(azure.recorded("scanned_p1"), same_keys=True)))
    result = azure.acquire(azure.scan_doc(tmp_path, 2), stub)
    facts = interpret(result.observations, result.coverage, DEFAULT_POLICIES)[1]
    return facts, TemplateExportAdapter().plan_bindings("s", tuple(facts))


def test_r7_4_page_two_depreciation_movements_and_prior_book_value(stub, tmp_path, monkeypatch):
    facts, _ = _run(stub, tmp_path, monkeypatch)
    page_two = collections.Counter(f.business.field for f in facts if f.source_pointer.get("page_index") == 1)
    assert not {"zugang", "abgang", "bw_close"} & set(page_two), page_two


def test_r7_6_one_asset_split_across_pages_with_disjoint_fields_binds_one_row(stub, tmp_path, monkeypatch):
    facts, bindings = _run(stub, tmp_path, monkeypatch)
    by_asset = collections.defaultdict(lambda: collections.defaultdict(set))
    for fact in facts:
        by_asset[fact.business.asset_id][fact.source_pointer.get("page_index")].add(fact.business.field)
    assert all(not (pages.get(0, set()) & pages.get(1, set())) for pages in by_asset.values()), "precondition: disjoint field sets"
    assert {page for pages in by_asset.values() for page in pages} == {0, 1}
    bound = [b for b in bindings if b.status == "bound"]
    # Fields without a template column stay contract_pending; every other binding binds.
    assert not [b for b in bindings if b.status not in {"bound", "contract_pending"}], collections.Counter(b.status for b in bindings)
    assert len(bound) >= 14
    assert len({b.destination.cell for b in bound}) == len(bound)
    rows = collections.defaultdict(set)
    for binding in bound:
        rows[binding.destination.row_key].add(binding.destination.cell.rsplit("!", 1)[1].lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
    assert all(len(numbers) == 1 for numbers in rows.values()), "one asset, one destination row"


def test_r7_6_overlapping_fields_on_two_pages_keep_refusing(stub, tmp_path, monkeypatch):
    monkeypatch.setenv(azure.KEY_ENV, "k")
    facts, bindings = r6._facts_and_bindings(stub, tmp_path, monkeypatch, same_keys=True)
    assert facts and not [b for b in bindings if b.status == "bound"]
