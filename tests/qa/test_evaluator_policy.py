"""Independent checks of the development evaluator's PDF row policy."""
from __future__ import annotations

import importlib.util
from hashlib import sha256
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("qa_evaluator", ROOT / "tests" / "acceptance" / "evaluate_corpus.py")
assert SPEC and SPEC.loader
EVALUATOR = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(EVALUATOR)


def fact_at(y: float, *, amount: str = "999", asset: str = "wrong") -> dict[str, object]:
    return {
        "field": "zugang",
        "pointer": {"kind": "pdf", "page_index": 0, "bbox": [40, y - 0.5, 60, y + 0.5]},
        "amount": {"normalized_base": amount},
        "asset_identity": {"asset_or_class": asset},
    }


def row(y0: float, y1: float, *, amount: str, asset: str) -> dict[str, object]:
    return {
        "field": "zugang",
        "pointer": {"page": 1, "bbox": [0, y0, 100, y1]},
        "amount": amount,
        "asset": asset,
    }


def test_unique_unexpanded_containment_resolves_tolerance_overlap_without_value_or_identity() -> None:
    expected = [row(0, 10, amount="1", asset="A"), row(11, 21, amount="999", asset="wrong")]
    # y=9.5 is in row 1's original box and in both boxes after the evaluator's
    # vertical tolerance expansion.  The deliberately matching value/identity
    # on row 2 must not influence selection.
    candidates = EVALUATOR.matching_candidates(fact_at(9.5), expected)
    assert [index for index, _ in candidates] == [0]


def test_no_unique_original_containment_remains_ambiguous_regardless_of_value_or_identity() -> None:
    expected = [row(0, 12, amount="1", asset="A"), row(10, 20, amount="999", asset="wrong")]
    first = EVALUATOR.matching_candidates(fact_at(11, amount="999", asset="wrong"), expected)
    second = EVALUATOR.matching_candidates(fact_at(11, amount="1", asset="A"), expected)
    assert [index for index, _ in first] == [0, 1]
    assert [index for index, _ in second] == [0, 1]


def xlsx_fact(
    fact_id: str,
    *,
    amount: str = "1",
    asset: str = "A-1",
    subnumber: str | None = None,
    period: str = "2025",
) -> dict[str, object]:
    return {
        "id": fact_id,
        "field": "zugang",
        "pointer": {"kind": "xlsx", "sheet": "Data", "cell": "B2"},
        "amount": {"normalized_base": amount, "unit": "EUR", "raw": amount},
        "asset_identity": {"asset_or_class": asset, "asset_sub_id": subnumber, "period": period},
    }


def evaluator_case(
    tmp_path: Path,
    *,
    rows: list[dict[str, object]],
    facts: list[dict[str, object]],
    abstentions: list[dict[str, object]] | None = None,
    unit: str = "EUR",
    status: str = "partial",
    complete_marker: str = "not-a-snapshot-digest",
) -> dict[str, object]:
    corpus = tmp_path / "corpus" / "f00_case"
    corpus.mkdir(parents=True)
    source = corpus / "case.xlsx"
    source.write_bytes(b"synthetic evaluator source")
    key = {"meta": {"unit": unit, "format": "xlsx", "family": "f00_case"}, "rows": rows}
    (corpus / "case.key.json").write_text(json.dumps(key))
    source_hash = sha256(source.read_bytes()).hexdigest()
    publication = tmp_path / "output" / "case-publication"
    publication.mkdir(parents=True)
    evidence = {
        "document": {"source_hash": source_hash, "status": status},
        "facts": facts,
        "abstentions": abstentions or [],
    }
    (publication / "evidence.v1.json").write_text(json.dumps(evidence))
    (publication / ".complete").write_text(complete_marker)
    return EVALUATOR.evaluate(tmp_path / "corpus", tmp_path / "output")


def key_row(
    *,
    cents: int = 100,
    printed: str | None = None,
    asset: str | None = "A-1",
    subnumber: str | None = None,
    year: int = 2025,
    class_name: str | None = None,
) -> dict[str, object]:
    value: dict[str, object] = {
        "fields": {"zugang": cents},
        "pointer": {"sheet": "Data", "cells": {"zugang": "B2"}},
        "asset_no": asset,
        "sub_no": subnumber,
        "year": year,
    }
    if printed is not None:
        value["printed"] = {"zugang": printed}
    if class_name is not None:
        value["class"] = class_name
    return value


def test_duplicate_acceptance_counts_only_first_fact_as_correct(tmp_path: Path) -> None:
    report = evaluator_case(tmp_path, rows=[key_row()], facts=[xlsx_fact("one"), xlsx_fact("two")])
    totals = report["totals"]
    assert totals["accepted"] == 2
    assert totals["correct"] == 1
    assert totals["wrong"] == 1
    assert totals["wrong_duplicate_acceptance"] == 1
    assert totals["unaccepted_slots"] == 0


def test_wrong_acceptance_and_abstention_counts_are_not_a_slot_partition(tmp_path: Path) -> None:
    report = evaluator_case(
        tmp_path,
        rows=[key_row()],
        facts=[xlsx_fact("wrong", amount="2")],
        abstentions=[{"id": "a"}, {"id": "b"}],
    )
    totals = report["totals"]
    assert totals["accepted"] == 1 and totals["wrong"] == 1
    assert totals["unaccepted_slots"] == 1
    assert totals["reported_abstentions"] == 2


def test_unevaluable_accepted_fact_is_excluded_from_precision_denominator(tmp_path: Path) -> None:
    report = evaluator_case(tmp_path, rows=[key_row(printed="not-a-number")], facts=[xlsx_fact("unevaluable")])
    totals = report["totals"]
    assert totals["source_slots"] == 1 and totals["accepted"] == 1
    assert totals["unevaluable_accepted"] == 1
    assert totals.get("correct", 0) == 0 and totals.get("wrong", 0) == 0
    assert totals["accepted_precision"] is None and totals["correct_coverage"] == 0


def test_proxy_does_not_enforce_class_empty_subnumber_exact_period_or_zero_unit(tmp_path: Path) -> None:
    report = evaluator_case(
        tmp_path,
        rows=[key_row(cents=0, printed="0", asset=None, subnumber=None, class_name="CLASS-A")],
        facts=[xlsx_fact("blind-spots", amount="0", asset="WRONG-CLASS", subnumber="SPURIOUS", period="garbage-2025-extra")],
        unit="TEUR",
    )
    assert report["totals"]["correct"] == 1, "documents a development-proxy limitation, not full-tuple correctness"


def test_complete_marker_contents_and_failed_status_do_not_gate_proxy_scoring(tmp_path: Path) -> None:
    report = evaluator_case(
        tmp_path,
        rows=[key_row()],
        facts=[xlsx_fact("accepted-despite-failed")],
        status="failed",
        complete_marker="tampered",
    )
    totals = report["totals"]
    assert totals["published_outputs"] == 1 and totals["status_failed"] == 1
    assert totals["correct"] == 1


def test_denominator_excludes_xlsx_field_without_a_physical_cell() -> None:
    key = {
        "meta": {"unit": "EUR"},
        "rows": [
            {
                "fields": {"zugang": 100, "abgang": 200},
                "pointer": {"sheet": "Data", "cells": {"zugang": "B2"}},
                "asset_no": "A-1",
                "year": 2025,
            }
        ],
    }
    expected = EVALUATOR.expected_values(key)
    assert [(item["field"], item["pointer"]["cells"].get(item["field"])) for item in expected] == [("zugang", "B2")]

