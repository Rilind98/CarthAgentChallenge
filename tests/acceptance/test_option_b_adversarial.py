"""Fix round 1: QA probes (qa_probes/) and reviewer repros (cr_exp/) as repo tests through the real run_input path.

Truth comes from the fixture builders' synthetic constants. A recorded plan may lose facts; it may never add a wrong one.
"""
from __future__ import annotations

import collections
import copy
from decimal import Decimal
import importlib.util
import json
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pytest

from carthagent.mapper import MapperConfig
from carthagent.pipeline import RunConfig, run_input

ROOT = Path(__file__).resolve().parents[2]


def _load(name: str):
    spec = importlib.util.spec_from_file_location(f"adv_{name}", ROOT / "tests" / "fixtures" / f"{name}.py")
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


fixtures = _load("build_option_b"); stub = _load("option_b_stub")
S = "Anlagengitter"
R = lambda n, sheet=S: f"@row:{sheet}:{n}"  # noqa: E731
C = lambda ref, sheet=S: f"@cell:{sheet}!{ref}"  # noqa: E731
SAP_FIELDS = (("ahk_open", "zugang", "abgang", "umb", None, None, "ahk_close"),
              ("kum_afa_open", "total_afa_gj", "afa_abgang", "afa_umb", None, "zuschreibung", "kum_afa_close"),
              ("bw_prev", None, None, None, None, None, "bw_close"))
REG_FIELDS = {5: "ahk_open", 6: "zugang", 7: "abgang", 8: "umb", 11: "ahk_close", 12: "kum_afa_open", 13: "total_afa_gj", 14: "afa_abgang", 15: "afa_umb",
              17: "zuschreibung", 18: "kum_afa_close", 19: "bw_prev", 20: "bw_close"}


def plan(name):
    return copy.deepcopy(stub.load_plan(name))


def run(tmp_path, source, recorded, out="out"):
    with stub.recorded_plan_server(recorded) as endpoint:
        result = run_input(source, tmp_path / out, RunConfig(mapper=MapperConfig(True, endpoint, "local-stub", 2, 256)))
    assert result.evidence_path is not None, result.diagnostics
    return result, json.loads(Path(result.evidence_path).read_text())["facts"]


def key(fact):
    return (fact["field"], fact["asset_identity"].get("asset_or_class"), fact["asset_identity"].get("period"), Decimal(fact["amount"]["normalized_base"]))


def sap_truth(blocks=None):
    truth = {}
    for offset, (identity, *measures) in enumerate(blocks or fixtures.SAP_BLOCKS):
        for index, (values, names) in enumerate(zip(measures, SAP_FIELDS), start=1):
            for column, value, field in zip("BCDEFGH", values, names):
                if field:
                    truth[(S, f"{column}{19 + offset * 5 + index}")] = (field, identity["E"], "2025", Decimal(str(value)))
    return truth


def reg_truth():
    return {("Sheet1", f"{get_column_letter(i + 1)}{14 + o}"): (f, v[0], "2025", Decimal(str(v[i]))) for o, v in enumerate(fixtures.REGISTER_ROWS) for i, f in REG_FIELDS.items()}


def spiegel_truth():
    names = ("ahk_open", "zugang", "abgang", "umb", "ahk_close", "kum_afa_open", "total_afa_gj", "afa_abgang", "kum_afa_close", "bw_close", "bw_prev")
    truth = {}
    for year, rows in fixtures.SPIEGEL_ROWS.items():
        for row, (number, name, *values) in rows.items():
            for column, value, field in zip(fixtures.SPIEGEL_VALUE_COLUMNS, values, names):
                truth[("Tabelle1" if year == 2025 else "Tabelle2", f"{fixtures.spiegel_column(column, year)}{row}")] = (field, name if year == 2025 else f"{number} {name}", str(year), Decimal(str(value)))
    return truth


def wrong(facts, truth):
    return [(f["pointer"]["sheet"], f["pointer"]["cell"], key(f)) for f in facts if truth.get((f["pointer"]["sheet"], f["pointer"]["cell"])) != key(f)]


def blk(identity, rows, role="data"):
    return {"identity_row_id": R(identity), "rows": [{"row_id": R(r), "header_row_id": (R(h) if h else None), "measure": m} for r, h, m in rows], "role": role}


def sap(tmp_path, recorded, class_cell=None):
    _, facts = run(tmp_path, fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx", class_cell=class_cell), recorded)
    return facts, wrong(facts, sap_truth())


def test_sap_baseline(tmp_path):
    facts, bad = sap(tmp_path, plan("sap_block_unlabelled"))
    assert len(facts) == 26 and not bad


SAP_VARIANTS = {
    "shift_minus_one": lambda q: (q.update(header_rows=[R(15), R(16), R(17)], field_headers=[f for f in q["field_headers"] if not f["observation_id"].endswith("18")]),
                                  q.update(record_blocks=[blk(18, [(19, 16, "ahk"), (20, 17, "cum_depreciation"), (21, None, None)]), blk(22, [(24, 16, "ahk"), (25, 17, "cum_depreciation"), (26, None, None)])])),
    "identity_points_at_numeric_row": lambda q: q.update(record_blocks=[blk(20, [(21, 16, "ahk"), (22, 17, "cum_depreciation")]), blk(25, [(26, 16, "ahk"), (27, 17, "cum_depreciation")])]),
    "swapped_headers_measure_null_two_rows": lambda q: q.update(record_blocks=[blk(19, [(20, 17, None), (21, 16, None)]), blk(24, [(25, 17, None), (26, 16, None)])]),
    "label_swap_zugang_umb": lambda q: [f.update(semantic_field={C("C16"): "umb", C("E16"): "zugang"}[f["observation_id"]]) for f in q["field_headers"] if f["observation_id"] in (C("C16"), C("E16"))],
    "label_swap_afa_abgang_zuschreibung": lambda q: [f.update(semantic_field={C("D17"): "zuschreibung", C("G17"): "afa_abgang"}[f["observation_id"]]) for f in q["field_headers"] if f["observation_id"] in (C("D17"), C("G17"))],
    "empty_identity_header_over_name_column": lambda q: q.update(identity_headers=[h for h in q["identity_headers"] if h["observation_id"] != C("E15")] + [{"observation_id": C("F15"), "identity_role": "class_code"}]),
    "wrong_text_identity_header": lambda q: q.update(identity_headers=[h for h in q["identity_headers"] if h["observation_id"] not in (C("E15"), C("D15"))] + [{"observation_id": C("D15"), "identity_role": "class_code"}]),
    "header_below": lambda q: (q.update(header_rows=[R(15), R(16), R(17), R(18), R(32)]), [b["rows"][0].update(header_row_id=R(32)) for b in q["record_blocks"]]),
    "header_two_fields": lambda q: q["field_headers"].append({"observation_id": C("B16"), "semantic_field": "zugang", "parent_header_ids": []}),
    "two_blocks_same_row": lambda q: q["record_blocks"].insert(1, blk(19, [(20, 16, "ahk"), (21, 17, "cum_depreciation"), (22, 18, "book_value")])),
    "non_periodic": lambda q: q["record_blocks"][1].update(rows=q["record_blocks"][1]["rows"][:2]),
    "span_widened": lambda q: q.update(column_spans=[{"header_observation_id": C("B16"), "first_column": 2, "last_column": 8}]),
    "span_inverted_and_outside": lambda q: q.update(column_spans=[{"header_observation_id": C("B16"), "first_column": 8, "last_column": 2}, {"header_observation_id": C("H16"), "first_column": 16000, "last_column": 99999999}, {"header_observation_id": C("H17"), "first_column": -5, "last_column": 0}]),
    "subtotal_declared_data": lambda q: q["record_blocks"][2].update(role="data"),
    "data_declared_subtotal": lambda q: q["record_blocks"][0].update(role="subtotal"),
}
SWAP_VARIANTS = {
    "3rows_measure_null_bv_governed": [(20, 17, None), (21, 16, None), (22, 18, None)],
    "3rows_measure_null_bv_ungoverned": [(20, 17, None), (21, 16, None), (22, None, None)],
    "2rows_measure_declared": [(20, 17, "ahk"), (21, 16, "cum_depreciation")],
    "2rows_measure_null": [(20, 17, None), (21, 16, None)],
    "2rows_measure_swapped_consistently": [(20, 17, "cum_depreciation"), (21, 16, "ahk")],
}


@pytest.mark.parametrize("name", list(SAP_VARIANTS))
def test_sap_adversarial_variant_binds_no_wrong_fact(tmp_path, name):
    recorded = plan("sap_block_unlabelled"); SAP_VARIANTS[name](recorded["plans"][0])
    facts, bad = sap(tmp_path, recorded)
    assert not bad, bad[:5]


@pytest.mark.parametrize("name", list(SWAP_VARIANTS))
def test_sap_header_row_swap_variant_binds_no_wrong_fact(tmp_path, name):
    rows = SWAP_VARIANTS[name]; recorded = plan("sap_block_unlabelled"); q = recorded["plans"][0]
    q["record_blocks"] = [blk(19, rows), blk(24, [(r + 5, h, m) for r, h, m in rows]), q["record_blocks"][2]]
    facts, bad = sap(tmp_path, recorded)
    assert not bad, bad[:5]


@pytest.mark.parametrize("afa_sign", [1, -1])
def test_consistent_ahk_afa_row_swap_with_positive_or_negative_afa_binds_nothing(tmp_path, afa_sign, monkeypatch):
    blocks = tuple((identity, ahk, tuple(None if v is None else (abs(v) if afa_sign > 0 else v) for v in afa), bw) for identity, ahk, afa, bw in fixtures.SAP_BLOCKS)
    monkeypatch.setattr(fixtures, "SAP_BLOCKS", blocks)
    recorded = plan("sap_block_unlabelled")
    for block in recorded["plans"][0]["record_blocks"]:
        rows = block["rows"]
        rows[0]["header_row_id"], rows[1]["header_row_id"] = rows[1]["header_row_id"], rows[0]["header_row_id"]
        rows[0]["measure"], rows[1]["measure"] = rows[1]["measure"], rows[0]["measure"]
    _, facts = run(tmp_path, fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), recorded)
    assert not wrong(facts, sap_truth(blocks))


def register(tmp_path, mutate=None):
    source = fixtures.flat_sap_register(tmp_path / "reg.xlsx")
    if mutate:
        workbook = load_workbook(source); mutate(workbook["Sheet1"]); workbook.save(source)
    return source


def test_register_baseline_and_shifted_identity(tmp_path):
    _, facts = run(tmp_path, register(tmp_path), plan("flat_sap_register"))
    assert facts and not wrong(facts, reg_truth())
    recorded = plan("flat_sap_register")
    recorded["plans"][0]["record_blocks"] = [{"identity_row_id": "@row:Sheet1:14", "rows": [{"row_id": "@row:Sheet1:15", "header_row_id": "@row:Sheet1:13", "measure": None}], "role": "data"},
                                            {"identity_row_id": "@row:Sheet1:16", "rows": [{"row_id": "@row:Sheet1:16", "header_row_id": "@row:Sheet1:13", "measure": None}], "role": "data"}]
    _, facts = run(tmp_path, register(tmp_path), recorded, "out2")
    assert not wrong(facts, reg_truth())


def test_register_empty_header_over_free_text_is_never_an_asset_id(tmp_path):
    def remark(ws):
        for row, text in ((14, "Umbuchung aus 100001002"), (15, "siehe Anlage 100001001"), (16, "Nachrüstung")):
            ws[f"V{row}"] = text
    recorded = plan("flat_sap_register"); q = recorded["plans"][0]
    q["identity_headers"] = [h for h in q["identity_headers"] if not h["observation_id"].endswith("A13")] + [{"observation_id": "@cell:Sheet1!V13", "identity_role": "asset_id"}]
    _, facts = run(tmp_path, register(tmp_path, remark), recorded)
    assert not wrong(facts, reg_truth())


@pytest.mark.parametrize("probe", ["capitalisation_date_block_scope", "amount_cell_sheet_scope"])
def test_register_period_never_comes_from_a_record_date_or_an_amount(tmp_path, probe):
    def mutate(ws):
        ws["A4"] = "Jahresabschluss"
        if probe == "amount_cell_sheet_scope":
            ws["G14"] = 2019
    recorded = plan("flat_sap_register"); q = recorded["plans"][0]
    q["period_contexts"] = ([{"observation_id": f"@cell:Sheet1!C{r}", "temporal_role": "closing", "scope": "block"} for r in (14, 15, 16)] if probe.startswith("capitalisation")
                            else [{"observation_id": "@cell:Sheet1!G14", "temporal_role": "closing", "scope": "sheet"}])
    _, facts = run(tmp_path, register(tmp_path, mutate), recorded)
    assert not facts


def test_register_two_year_cell_and_unit_conflict(tmp_path):
    def mutate(ws):
        ws["V14"] = "GJ 2024/2025"; ws["W14"] = "TEUR"
    recorded = plan("flat_sap_register"); q = recorded["plans"][0]
    q["period_contexts"] = [{"observation_id": "@cell:Sheet1!V14", "temporal_role": "closing", "scope": "block"}]
    q["unit_contexts"] = [{"observation_id": "@cell:Sheet1!E15", "unit": "EUR", "scope": "sheet"}, {"observation_id": "@cell:Sheet1!W14", "unit": "TEUR", "scope": "block"}]
    _, facts = run(tmp_path, register(tmp_path, mutate), recorded)
    assert not wrong(facts, reg_truth())


def test_spiegel_baseline_foreign_header_row_and_measure_swap(tmp_path):
    source = fixtures.two_year_two_sheets(tmp_path / "s.xlsx")
    _, facts = run(tmp_path, source, plan("two_year_two_sheets"))
    assert facts and not wrong(facts, spiegel_truth())
    foreign = plan("two_year_two_sheets"); q = foreign["plans"][0]
    q["header_rows"].append("@row:Tabelle2:5")
    for block in q["record_blocks"]:
        block["rows"][0]["header_row_id"] = "@row:Tabelle2:5"
    _, facts = run(tmp_path, source, foreign, "out2")
    assert not wrong(facts, spiegel_truth()) and not [f for f in facts if f["pointer"]["sheet"] == "Tabelle1"]
    swapped = plan("two_year_two_sheets")
    for q in swapped["plans"]:
        for axis in q["measure_axes"]:
            axis["semantic_measure"] = {"ahk": "cum_depreciation", "cum_depreciation": "ahk"}.get(axis["semantic_measure"], axis["semantic_measure"])
    _, facts = run(tmp_path, source, swapped, "out3")
    assert not wrong(facts, spiegel_truth())


@pytest.mark.parametrize("role", ["asset_id", "class_code", "account", "entity"])
def test_empty_label_column_header_carries_only_descriptive_roles(tmp_path, role):
    recorded = plan("two_year_two_sheets")
    for q in recorded["plans"]:
        for header in q["identity_headers"]:
            header["identity_role"] = role
    _, facts = run(tmp_path, fixtures.two_year_two_sheets(tmp_path / "s.xlsx"), recorded)
    assert not facts


def test_undeclared_sub_number_header_keeps_the_local_abstentions(tmp_path):
    def mutate(ws):
        ws["B13"] = "Unternummer"; ws["B15"] = None; ws["B16"] = None
    local_run, local_facts = run(tmp_path, register(tmp_path, mutate), {"plans": []}, "local")
    recorded = plan("flat_sap_register"); q = recorded["plans"][0]
    q["identity_headers"] = [h for h in q["identity_headers"] if h["identity_role"] != "asset_sub_id"]
    _, facts = run(tmp_path, register(tmp_path, mutate), recorded, "mapped")
    assert {(f["pointer"]["cell"], key(f)) for f in facts} == {(f["pointer"]["cell"], key(f)) for f in local_facts}


@pytest.mark.parametrize("header", ["UNr", "Unter-Nr.", "lfd. Nr."])
def test_undeclared_identifier_column_in_header_area_skips_the_block(tmp_path, header):
    source = fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx")
    workbook = load_workbook(source); ws = workbook.active
    ws["C15"] = header; ws["C19"] = "0"; ws["C24"] = "1"; ws["E24"] = "1000"; workbook.save(source)
    recorded = plan("sap_block_unlabelled"); q = recorded["plans"][0]
    q["identity_headers"] = [h for h in q["identity_headers"] if not h["observation_id"].endswith("!C15")]
    _, facts = run(tmp_path, source, recorded)
    keys = collections.Counter((f["field"], f["asset_identity"]["asset_or_class"], f["asset_identity"]["asset_sub_id"]) for f in facts)
    assert not facts and not [k for k, n in keys.items() if n > 1]
