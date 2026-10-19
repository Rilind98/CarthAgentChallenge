"""Round 6 item 6: the same asset keys on two scanned pages survive semantics (K1 is per page) and the exporter refuses
every destination they would share, end to end (layout engine -> semantics -> binding plan)."""

from __future__ import annotations

import collections
import copy
import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


azure = _load(ROOT / "tests" / "unit" / "test_azure_layout.py", "r6_azure_layout")
stub = azure.stub  # noqa: F401
documents = azure.documents  # noqa: F401


def _two_pages(payload, same_keys):
    """Copy page 1 and its table onto page 2 (QA round-5 probe); optionally renumber the assets on page 2."""
    result = payload["analyzeResult"]; page_two = copy.deepcopy(result["pages"][0]); page_two["pageNumber"] = 2
    content = result["content"]; moved = {}
    renumber = {"000101": "000303", "000202": "000404"}
    for word in page_two["words"]:
        offset = len(content) + 1; content += " " + word["content"]; moved[word["span"]["offset"]] = offset
        word["span"] = {"offset": offset, "length": len(word["content"])}
        if not same_keys and word["content"] in renumber:
            word["content"] = renumber[word["content"]]
    table = copy.deepcopy(result["tables"][0])
    for cell in table["cells"]:
        for region in cell["boundingRegions"]:
            region["pageNumber"] = 2
        cell["spans"] = [{"offset": moved.get(span["offset"], span["offset"]), "length": span["length"]} for span in cell.get("spans", [])]
        if not same_keys and cell["content"] in renumber:
            cell["content"] = renumber[cell["content"]]
    for region in table["boundingRegions"]:
        region["pageNumber"] = 2
    result["pages"].append(page_two); result["tables"].append(table); result["content"] = content
    return payload


def _facts_and_bindings(stub, tmp_path, monkeypatch, same_keys):
    from carthagent.exporter import TemplateExportAdapter
    from carthagent.policies import DEFAULT_POLICIES
    from carthagent.semantics import interpret
    monkeypatch.setenv(azure.KEY_ENV, "k"); azure.LayoutStub.reset(_two_pages(azure.recorded("scanned_p1"), same_keys))
    result = azure.acquire(azure.scan_doc(tmp_path, 2), stub)
    facts = interpret(result.observations, result.coverage, DEFAULT_POLICIES)[1]
    return facts, TemplateExportAdapter().plan_bindings("s", tuple(facts))


def test_item6_same_keys_on_two_pages_never_bind_one_destination(stub, tmp_path, monkeypatch):
    facts, bindings = _facts_and_bindings(stub, tmp_path, monkeypatch, same_keys=True)
    keys = collections.defaultdict(set)
    for fact in facts:
        keys[(fact.business.asset_or_class, fact.business.field, fact.business.period)].add(fact.source_pointer.get("page_index"))
    shared = {key for key, pages in keys.items() if len(pages) > 1}
    assert shared, "precondition: K1 is per page, so the same keys reach the exporter from both pages"
    by_fact = {binding.record_id: binding for binding in bindings}
    for fact in facts:
        if (fact.business.asset_or_class, fact.business.field, fact.business.period) in shared:
            binding = by_fact[fact.id]
            assert binding.status != "bound" and binding.destination.cell is None, (fact.business.asset_or_class, fact.business.field, binding.status)


def test_item6_continuation_page_with_new_keys_binds_both_pages(stub, tmp_path, monkeypatch):
    facts, bindings = _facts_and_bindings(stub, tmp_path, monkeypatch, same_keys=False)
    assert {fact.source_pointer.get("page_index") for fact in facts} == {0, 1}
    bound = [binding for binding in bindings if binding.status == "bound"]
    assert len(bound) == len(facts) and len({b.destination.cell for b in bound}) == len(bound)
