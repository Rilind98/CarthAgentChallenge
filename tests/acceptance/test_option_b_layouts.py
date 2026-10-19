"""Option B acceptance: recorded model plans declare structure; code verifies it before binding any value.

Expected facts are derived from the fixture builders' synthetic constants, not from the code under test.
Adversarial recorded plans must yield zero facts for the records they corrupt.
"""
from __future__ import annotations

from decimal import Decimal
import importlib.util
import json
from pathlib import Path

import pytest

from carthagent.mapper import MapperConfig
from carthagent.pipeline import RunConfig, run_input

ROOT = Path(__file__).resolve().parents[2]


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / "tests" / "fixtures" / f"{name}.py")
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


fixtures = _load("build_option_b"); stub = _load("option_b_stub")
SAP_FIELDS = (("ahk_open", "zugang", "abgang", "umb", None, None, "ahk_close"),
              ("kum_afa_open", "total_afa_gj", "afa_abgang", "afa_umb", None, "zuschreibung", "kum_afa_close"),
              ("bw_prev", None, None, None, None, None, "bw_close"))
SPIEGEL_FIELDS = ("ahk_open", "zugang", "abgang", "umb", "ahk_close", "kum_afa_open", "total_afa_gj", "afa_abgang", "kum_afa_close", "bw_close", "bw_prev")
REGISTER_FIELDS = {5: "ahk_open", 6: "zugang", 7: "abgang", 8: "umb", 11: "ahk_close", 12: "kum_afa_open", 13: "total_afa_gj", 14: "afa_abgang", 15: "afa_umb",
                   17: "zuschreibung", 18: "kum_afa_close", 19: "bw_prev", 20: "bw_close"}


# Local header rules once read "AfA GJ-Beg" as total_afa_gj (f09) and "Zugänge" under the AfA group as zugang (f08); both fixed,
# so local and block agree everywhere. A disagreement would still make the value abstain (second opinion).
# The verified block disagrees, so those values must abstain (second opinion), never take either label.
LOCAL_DISAGREES: dict[str, tuple[str, ...]] = {}


def expected_sap() -> set[tuple]:
    result = set()
    for offset, (identity, *measures) in enumerate(fixtures.SAP_BLOCKS):
        for index, (values, names) in enumerate(zip(measures, SAP_FIELDS), start=1):
            for column, value, field in zip("BCDEFGH", values, names):
                if field:
                    result.add(("Anlagengitter", f"{column}{19 + offset * 5 + index}", field, identity["E"], "2025", Decimal(str(value))))
    return result


def expected_spiegel() -> set[tuple]:
    result = set()
    for year, rows in fixtures.SPIEGEL_ROWS.items():
        sheet = "Tabelle1" if year == 2025 else "Tabelle2"
        for row, (number, name, *values) in rows.items():
            for column, value, field in zip(fixtures.SPIEGEL_VALUE_COLUMNS, values, SPIEGEL_FIELDS):
                result.add((sheet, f"{fixtures.spiegel_column(column, year)}{row}", field, name if year == 2025 else f"{number} {name}", str(year), Decimal(str(value))))
    return result


def expected_register() -> set[tuple]:
    from openpyxl.utils import get_column_letter
    return {("Sheet1", f"{get_column_letter(index + 1)}{14 + offset}", field, values[0], "2025", Decimal(str(values[index])))
            for offset, values in enumerate(fixtures.REGISTER_ROWS) for index, field in REGISTER_FIELDS.items()}


def run(tmp_path: Path, source: Path, plan: dict) -> tuple[dict, list[dict]]:
    with stub.recorded_plan_server(plan) as endpoint:
        result = run_input(source, tmp_path / "out", RunConfig(mapper=MapperConfig(True, endpoint, "local-stub", 2, 256)))
    assert result.evidence_path is not None, result.diagnostics
    assert not [item for item in result.diagnostics if item.startswith("mapper_unavailable")], result.diagnostics
    evidence = json.loads(Path(result.evidence_path).read_text())
    return evidence, evidence["facts"]


def facts_as_rows(facts: list[dict]) -> set[tuple]:
    return {(f["pointer"]["sheet"], f["pointer"]["cell"], f["field"], f["asset_identity"]["asset_or_class"], f["asset_identity"]["period"], Decimal(f["amount"]["normalized_base"])) for f in facts}


@pytest.mark.parametrize(("builder", "plan", "expected"), [
    ("sap_block_unlabelled", "sap_block_unlabelled", expected_sap),
    ("two_year_two_sheets", "two_year_two_sheets", expected_spiegel),
    ("flat_sap_register", "flat_sap_register", expected_register),
])
def test_recorded_structure_plan_binds_exactly_the_expected_facts(tmp_path, builder, plan, expected):
    source = getattr(fixtures, builder)(tmp_path / f"{builder}.xlsx")
    _, facts = run(tmp_path, source, stub.load_plan(plan))
    assert facts_as_rows(facts) == {row for row in expected() if row[2] not in LOCAL_DISAGREES.get(builder, ())}
    assert facts and all(f["amount"]["unit"] == "EUR" for f in facts)


def test_f05_shape_one_metric_cells(tmp_path):
    _, facts = run(tmp_path, fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), stub.load_plan("sap_block_unlabelled"))
    by_cell = {f["pointer"]["cell"]: f for f in facts}
    for cell, field in (("B20", "ahk_open"), ("B21", "kum_afa_open"), ("H22", "bw_close")):
        identity = by_cell[cell]["asset_identity"]
        assert (by_cell[cell]["field"], identity["class_code"], identity["period"], by_cell[cell]["amount"]["unit"], by_cell[cell]["pointer"]["sheet"]) == (field, "1000", "2025", "EUR", "Anlagengitter")


@pytest.mark.parametrize(("plan", "class_cell"), [
    ("adversarial_shifted_block", None),
    ("adversarial_swapped_header_rows", None),
    ("adversarial_numeric_identity", 1000.5),
])
def test_adversarial_recorded_plans_yield_zero_facts_for_the_corrupted_records(tmp_path, plan, class_cell):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx", class_cell=class_cell)
    _, facts = run(tmp_path, source, stub.load_plan(plan))
    corrupted = {f"{column}{row}" for column in "BCDEFGH" for row in (20, 21, 22)}
    if plan != "adversarial_numeric_identity":
        corrupted |= {f"{column}{row}" for column in "BCDEFGH" for row in (25, 26, 27)}
    assert not [f["pointer"]["cell"] for f in facts if f["pointer"]["cell"] in corrupted]
    if plan != "adversarial_numeric_identity":
        assert facts == []
