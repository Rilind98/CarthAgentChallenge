"""Fix round 3: re-QA round-2 probes (qa_probes/test_round2_*.py, dbg_grp.py, round2_local_vs_mapper.py, local_unr.py)."""
from __future__ import annotations

import collections
from decimal import Decimal
import importlib.util
import json
from pathlib import Path

import pytest

from carthagent.mapper import _supported_roles
from carthagent.pipeline import RunConfig, run_input

ROOT = Path(__file__).resolve().parents[2]


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


adv = _load(Path(__file__).with_name("test_option_b_adversarial.py"), "r3_adversarial")
split = _load(ROOT / "tests" / "fixtures" / "option_b_split_stub.py", "r3_split")


def ident(*pairs):
    return [{"observation_id": f"@cell:Sheet1!{c}", "identity_role": r} for c, r in pairs]


def run_register(tmp_path, mutate, ids=None, periods=None):
    """(mapper-added wrong facts, colliding business keys) for a register variant; local-only facts are subtracted."""
    source = adv.register(tmp_path, mutate)
    local = json.loads(Path(run_input(source, tmp_path / "local", RunConfig()).evidence_path).read_text())["facts"]
    recorded = adv.plan("flat_sap_register"); q = recorded["plans"][0]
    if ids is not None:
        q["identity_headers"] = ids
    if periods is not None:
        q["period_contexts"] = periods
    _, facts = adv.run(tmp_path, source, recorded, "mapped")
    identity = lambda f: (f["pointer"]["cell"], f["field"], json.dumps(f["asset_identity"], sort_keys=True), f["amount"]["normalized_base"])  # noqa: E731
    local_keys = {identity(f) for f in local}
    truth = adv.reg_truth()
    added_wrong = [f for f in facts if identity(f) not in local_keys and truth.get((f["pointer"]["sheet"], f["pointer"]["cell"])) != adv.key(f)]
    keys = collections.Counter((f["asset_identity"].get("asset_or_class"), f["asset_identity"].get("asset_sub_id"), f["field"], f["asset_identity"].get("period")) for f in facts)
    return added_wrong, [k for k, n in keys.items() if n > 1]


# --- R3-1 / R3-2: periods ---

@pytest.mark.parametrize("stamp", ["Stand: 15.03.2026", "Datum 26.04.2026", "Seite 1 per 26.04.2026", "Abfrage zum 01.02.2026", "Gültig ab 01.01.2019", "Erstellt am 02.02.2026"])
def test_sheet_print_stamps_are_never_the_reporting_period(tmp_path, stamp):
    def mutate(ws):
        ws["A4"] = "Jahresabschluss"; ws["A6"] = stamp
    bad, _ = run_register(tmp_path, mutate, periods=[{"observation_id": "@cell:Sheet1!A6", "temporal_role": "closing", "scope": "sheet"}])
    assert not bad


@pytest.mark.parametrize("reporting", ["Jahresabschluss", "Jahresabschluss per 30.06.2025"])
def test_block_period_never_from_a_name_cell_with_period_words(tmp_path, reporting):
    def mutate(ws):
        ws["A4"] = reporting; ws["D14"] = "Zugang Geschäftsjahr 2023"; ws["D15"] = "Umbau Wirtschaftsjahr 2022"; ws["D16"] = "Erweiterung GJ 2021"
    bad, _ = run_register(tmp_path, mutate, ident(("A13", "asset_id"), ("B13", "asset_sub_id")),
                          [{"observation_id": f"@cell:Sheet1!D{r}", "temporal_role": "closing", "scope": "block"} for r in (14, 15, 16)])
    assert not bad


def test_block_period_never_from_an_undeclared_comment_column(tmp_path):
    def mutate(ws):
        ws["A4"] = "Jahresabschluss"; ws["V13"] = "Bemerkung"
        for row, text in ((14, "Zugang im Geschäftsjahr 2019"), (15, "Stand 2020"), (16, "per 2021 umgebucht")):
            ws[f"V{row}"] = text
    bad, _ = run_register(tmp_path, mutate, periods=[{"observation_id": f"@cell:Sheet1!V{r}", "temporal_role": "closing", "scope": "block"} for r in (14, 15, 16)])
    assert not bad


# --- R3-3: asset_id support ---

def test_supported_roles_are_anchored():
    assert {t: sorted(_supported_roles(t)) for t in ("Anlagengruppe", "Anlagedatum", "Inventurdatum", "Objektart", "Anlagen im Bau", "Unternehmen", "lfd Buchwert",
                                                     "Gegenkonto", "Teilwert")} == {t: [] for t in ("Anlagengruppe", "Anlagedatum", "Inventurdatum", "Objektart", "Anlagen im Bau",
                                                                                                    "Unternehmen", "lfd Buchwert", "Gegenkonto", "Teilwert")} | {"Gegenkonto": ["account"]}
    for text, role in (("Anlage", "asset_id"), ("Anlagennummer", "asset_id"), ("Anl.-Nr.", "asset_id"), ("Inventarnummer", "asset_id"), ("Inv.-Nr.", "asset_id"),
                       ("Asset ID", "asset_id"), ("Klasse", "class_code"), ("Unter-nummer", "asset_sub_id"), ("UNr", "asset_sub_id"), ("Konto", "account")):
        assert role in _supported_roles(text), text


@pytest.mark.parametrize(("a13", "c13"), [("Kennung", "Anlagengruppe"), ("Kennung", "Anlagedatum"), ("Kennung-X", "Anlagengruppe"), ("ID", "Anlagengruppe"), ("Ident", "Anlagengruppe")])
def test_lookalike_id_header_with_an_undeclared_true_id_column_binds_nothing_new(tmp_path, a13, c13):
    def mutate(ws):
        ws["A13"] = a13; ws["A6"] = None; ws["C13"] = c13; ws["D13"] = "Bezeichnung"
        for row in (14, 15, 16):
            ws[f"C{row}"] = "0027"
    bad, collisions = run_register(tmp_path, mutate, ident(("C13", "asset_id"), ("B13", "asset_sub_id"), ("D13", "source_name")))
    assert not bad and not collisions


def test_nummer_alone_and_cost_centre_konto(tmp_path):
    bad, collisions = run_register(tmp_path, lambda ws: ws.__setitem__("A13", "Nummer"), ident(("A13", "asset_id"), ("B13", "asset_sub_id"), ("D13", "source_name")))
    assert not bad and not collisions

    def cost_centre(ws):
        ws["V13"] = "Kst.-Konto"
        for row, value in ((14, "4711"), (15, "4712"), (16, "4713")):
            ws[f"V{row}"] = value
    (tmp_path / "kst").mkdir()
    bad, collisions = run_register(tmp_path / "kst", cost_centre, ident(("A13", "asset_id"), ("B13", "asset_sub_id"), ("D13", "source_name"), ("V13", "account")))
    assert not bad and not collisions


# --- R3-4: header band carry with two tables on one sheet ---

@pytest.mark.parametrize(("a_blocks", "abbreviated"), [(2, False), (8, False), (8, True)])
def test_split_requests_never_govern_a_table_with_another_tables_header_band(tmp_path, a_blocks, abbreviated):
    source, band_b, truth = split.build_two_tables(tmp_path / "two.xlsx", a_blocks, abbreviated)
    split.NearestBand.bands = (15, band_b); split.NearestBand.pieces = []
    result = split.serve(split.NearestBand, source, tmp_path / "o")
    facts = json.loads(Path(result.evidence_path).read_text())["facts"]
    assert not adv.wrong(facts, truth)


def test_a_header_row_across_another_tables_band_is_refused(tmp_path):
    """Both bands visible in one request; the plan governs table B by table A's header rows: R1 refuses it."""
    source, band_b, truth = split.build_two_tables(tmp_path / "two.xlsx", 2, False, b_blocks=4)

    class OtherBand(split.NearestBand):
        pick = staticmethod(min)
    OtherBand.bands = (15, band_b); OtherBand.pieces = []
    result = split.serve(OtherBand, source, tmp_path / "o", chunk=10_000)
    facts = json.loads(Path(result.evidence_path).read_text())["facts"]
    assert len(OtherBand.pieces) == 1 and OtherBand.pieces[0]["bands"] == [15, band_b]
    assert not adv.wrong(facts, truth) and not [f for f in facts if int(f["pointer"]["cell"][1:]) > band_b]


# --- R3-5: local path with sub-number header variants (R2-8) ---

REG_FIELDS = adv.REG_FIELDS


@pytest.mark.parametrize("text", ["UNr", "U-Nr.", "Sub", "Teil", "Lfd.", "Unterpos.", "SNr", "Unter-nummer"])
def test_local_rules_never_key_by_name_or_read_afa_gj_beg_as_charge(tmp_path, text):
    source = adv.register(tmp_path, lambda ws: ws.__setitem__("B13", text))
    facts = json.loads(Path(run_input(source, tmp_path / "o", RunConfig()).evidence_path).read_text())["facts"]
    assert not adv.wrong(facts, adv.reg_truth())
    assert not [f for f in facts if f["pointer"]["cell"].startswith("M") and f["field"] == "total_afa_gj"]


@pytest.mark.parametrize("case", ["unknown_id_header", "unknown_sub_number_header"])
def test_local_rules_abstain_when_an_unknown_identifier_header_holds_values(tmp_path, case):
    """After R3-5 the local rules know "Anlage"/"Unter-nummer"; an unknown id-like header ("ID", "Kz") must still make
    them abstain instead of keying by name or merging two sub-numbered records under one key."""
    def mutate(ws):
        if case == "unknown_id_header":
            ws["A13"] = "ID"
        else:
            ws["B13"] = "Kz"
    source = adv.register(tmp_path, mutate)
    facts = json.loads(Path(run_input(source, tmp_path / "o", RunConfig()).evidence_path).read_text())["facts"]
    keys = collections.Counter((f["asset_identity"].get("asset_or_class"), f["asset_identity"].get("asset_sub_id"), f["field"]) for f in facts)
    assert not adv.wrong(facts, adv.reg_truth()) and not [k for k, n in keys.items() if n > 1]


# --- 3.1-b: reconciliation sections below a table are never records (local rules and mapper) ---

def _below_table(facts):
    return [f["pointer"]["cell"] for f in facts if f["pointer"]["sheet"] == "Tabelle1" and int(f["pointer"]["cell"].lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")) >= 19]


def test_reconciliation_section_binds_nothing_offline(tmp_path):
    source = adv.fixtures.two_year_two_sheets(tmp_path / "s.xlsx")
    for name, recorded in (("local", {"plans": []}), ("mapped", adv.plan("two_year_two_sheets"))):
        _, facts = adv.run(tmp_path, source, recorded, name)
        assert not _below_table(facts), (name, _below_table(facts))
        assert not adv.wrong(facts, adv.spiegel_truth()), name
    assert len(facts) == 66, "every main-table value still binds"


def test_plan_declared_reconciliation_rows_as_data_still_bind_nothing(tmp_path):
    recorded = adv.plan("two_year_two_sheets"); q = recorded["plans"][0]
    q["record_blocks"] += [{"identity_row_id": f"@row:Tabelle1:{r}", "rows": [{"row_id": f"@row:Tabelle1:{r}", "header_row_id": "@row:Tabelle1:5", "measure": None}], "role": "data"}
                           for r in (22, 23, 26, 27, 28)]
    _, facts = adv.run(tmp_path, adv.fixtures.two_year_two_sheets(tmp_path / "s.xlsx"), recorded)
    assert not _below_table(facts) and not adv.wrong(facts, adv.spiegel_truth())


F08 = ROOT / "hackathon-local" / "ElabHackathon" / "A3-Hackathon-Corpus" / "f08_excel_spiegel_2y" / "f08_excel_spiegel_2y_01.xlsx"


@pytest.mark.skipif(not F08.exists(), reason="private corpus unavailable")
def test_live_stage1_f08_plans_bind_nothing_from_the_reconciliation_rows():
    from carthagent.acquire_xlsx import acquire_xlsx
    from carthagent.mapper import MapperConfig, OpenAICompatibleMapper
    from carthagent.policies import DEFAULT_BOUNDS, DEFAULT_POLICIES
    from carthagent.semantics import interpret
    records = [json.loads((ROOT / "tests" / "fixtures" / "option_b_plans" / f"live_stage1_f08_tabelle{n}.json").read_text()) for n in (1, 2)]
    document_id = records[0]["document_id"]
    observations = acquire_xlsx(F08, document_id, "a" * 64, DEFAULT_BOUNDS).observations
    mapper = OpenAICompatibleMapper(MapperConfig(True, "http://127.0.0.1:1/v1/responses", "replay", 1, 1)); mapper.last_verification = []
    mapped = mapper._validate(document_id, observations, {"plans": [p for record in records for p in record["plans"]]})
    facts = interpret(observations, (), DEFAULT_POLICIES, mapped)[1]
    rows = [(f.physical.occurrence.split("!")[0], int(f.physical.occurrence.split("!")[1].lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))) for f in facts]
    assert not [r for r in rows if r[0] == "Tabelle1" and 29 <= r[1] <= 39], rows
    assert len([r for r in rows if 10 <= r[1] <= 24]) >= 25, rows
