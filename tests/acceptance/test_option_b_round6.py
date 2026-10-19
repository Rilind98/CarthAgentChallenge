"""Round 6: regressions seen in live stage-1 on 935f331, replayed offline from the recorded live plans.

R6-1 f00_demo_01 local-only must equal the 779e1a4 facts (no new abstentions).
R6-2 the live f05 plan (F16/G16 labelled bew_ahk under AHK) must bind through run_input with 0 wrong vs the sheet.
R6-3 the live f08 plans must bind at least 25 main-table cells with 0 wrong and 0 reconciliation rows.
"""

from __future__ import annotations

from collections import Counter
from decimal import Decimal
import importlib.util
import json
import re
from pathlib import Path

import pytest

from carthagent.mapper import MapperConfig
from carthagent.pipeline import RunConfig, run_input

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "hackathon-local" / "ElabHackathon" / "A3-Hackathon-Corpus"
PLANS = ROOT / "tests" / "fixtures" / "option_b_plans"
F00 = CORPUS / "f00_demo" / "f00_demo_01.xlsx"
F05 = CORPUS / ("f05_sap_class_" "gitter") / ("f05_sap_class_" "gitter_02.xlsx")


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


stub = _load(ROOT / "tests" / "fixtures" / "option_b_stub.py", "r6_stub")


def replay(source: Path, name: str, out: Path) -> dict:
    """Run the real pipeline with both recorded live requests served by the loopback stub (plan IDs untouched)."""
    plans = [plan for n in (1, 2) for plan in json.loads((PLANS / f"live_r5_{name}_req{n}.json").read_text())["plans"]]
    with stub.recorded_plan_server({"plans": plans}) as endpoint:
        result = run_input(source, out, RunConfig(mapper=MapperConfig(True, endpoint, "replay", 2, 256)))
    assert not [d for d in result.diagnostics if "mapper" in d and "error" in d.lower()], result.diagnostics
    return json.loads(Path(result.evidence_path).read_text())


F00_EXPECTED = ROOT / "hackathon-local" / "test-fixtures" / "f00_demo_01_local_779e1a4.json"


@pytest.mark.skipif(not (F00.exists() and F00_EXPECTED.exists()), reason="private corpus or its recorded local facts unavailable")
def test_r6_1_f00_local_only_equals_779e1a4():
    from carthagent.acquire_xlsx import acquire_xlsx
    from carthagent.policies import DEFAULT_BOUNDS, DEFAULT_POLICIES
    from carthagent.semantics import interpret
    expected = {tuple(row) for row in json.loads(F00_EXPECTED.read_text())["facts"]}
    observations = acquire_xlsx(F00, "d", "a" * 64, DEFAULT_BOUNDS).observations
    _, facts, abstentions, _ = interpret(observations, (), DEFAULT_POLICIES)
    got = {(f.physical.occurrence, f.business.field, str(f.amount.normalized_value), f.business.asset_or_class, f.business.period) for f in facts}
    assert len(expected) == 415
    assert got == expected, (sorted(expected - got)[:5], sorted(got - expected)[:5])
    assert not abstentions, Counter(a.reason for a in abstentions)


@pytest.mark.skipif(not F05.exists(), reason="private corpus unavailable")
def test_r6_2_live_f05_plan_binds_through_run_input_with_zero_wrong(tmp_path):
    qa4 = _load(Path(__file__).with_name("test_option_b_round5_qa4.py"), "r6_qa4")
    evidence = replay(F05, "f05_sap_class_gitter_02", tmp_path / "out")
    truth = qa4.sheet_truth(F05)
    got = {f["pointer"]["cell"]: (f["field"], f["asset_identity"].get("class_code"), Decimal(str(f["amount"]["normalized_base"])), f["asset_identity"].get("period"), f["amount"]["unit"])
           for f in evidence["facts"] if f["pointer"]["sheet"] == "Anlagengitter"}
    assert got, "the live plan binds"
    wrong = [(cell, value) for cell, value in got.items() if truth.get(cell) != value[:3] or value[3:] != ("2025", "EUR")]
    assert not wrong, wrong[:5]
    assert got["B20"] == ("ahk_open", "1000", Decimal("208034.23"), "2025", "EUR")
    assert got["B21"][:2] == ("kum_afa_open", "1000") and got["H22"][:2] == ("bw_close", "1000")
    assert not {cell for cell in got if cell[0] in "FG" and int(cell[1:]) % 5 == 0}, "Nachaktivg / Invest.Foerderg (labelled bew_ahk under AHK) abstain"


F08 = CORPUS / "f08_excel_spiegel_2y" / "f08_excel_spiegel_2y_01.xlsx"
SPIEGEL_FIELDS = ("ahk_open", "zugang", "abgang", "umb", "ahk_close", "kum_afa_open", "total_afa_gj", "afa_abgang", "kum_afa_close", "bw_close", "bw_prev")
F08_COLUMNS = {"Tabelle1": dict(zip("CDEFHJKLMOP", SPIEGEL_FIELDS)), "Tabelle2": dict(zip("BCDEGIJKLNO", SPIEGEL_FIELDS))}


def f08_sheet_truth(path):
    """Truth read from the sheet: numbered class rows ("2." in A or leading the name), field from the column, empty = 0."""
    from openpyxl import load_workbook
    truth = {}
    workbook = load_workbook(path, data_only=True)
    for sheet, (name_column, period) in {"Tabelle1": ("B", "2025"), "Tabelle2": ("A", "2024")}.items():
        ws = workbook[sheet]
        for row in range(7, 27):
            if not re.match(r"\d+\.", str(ws[f"A{row}"].value or "").strip()) or not ws[f"{name_column}{row}"].value:
                continue
            for column, field in F08_COLUMNS[sheet].items():
                value = ws[f"{column}{row}"].value
                truth[(sheet, f"{column}{row}")] = (field, Decimal(str(value)) if isinstance(value, (int, float)) else Decimal(0), period, "EUR")
    return truth


@pytest.mark.skipif(not F08.exists(), reason="private corpus unavailable")
def test_r6_3_live_f08_plans_bind_main_table_with_units_and_zero_wrong(tmp_path):
    evidence = replay(F08, "f08_excel_spiegel_2y_01", tmp_path / "out")
    truth = f08_sheet_truth(F08)
    got = {(f["pointer"]["sheet"], f["pointer"]["cell"]): (f["field"], Decimal(str(f["amount"]["normalized_base"])), f["asset_identity"].get("period"), f["amount"]["unit"]) for f in evidence["facts"]}
    assert not [key for key in got if int(key[1][1:]) >= 27], "reconciliation rows below the table never bind"
    wrong = [(key, value, truth.get(key)) for key, value in got.items() if truth.get(key) != value]
    assert not wrong, wrong[:5]
    correct = sum(1 for key, value in got.items() if truth.get(key) == value)
    assert correct >= 25, (correct, Counter(a["reason"] for a in evidence["abstentions"]).most_common(4))
    assert got[("Tabelle1", "C10")] == ("ahk_open", Decimal("144589.68"), "2025", "EUR")
