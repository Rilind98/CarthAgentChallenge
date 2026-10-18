"""f06_sap_asset_gitter: SAP-like Anlagengitter, one row per asset (UNr. = sub-number), 16 boxed columns; the landscape
list is turned 90° counter-clockwise onto portrait pages. A block under the page header names Buchungskreis,
Bilanzposition, Bestandskonto and class; new page at each class; only a '*Klasse' subtotal, no grand total;
descriptions cut at a fixed width; selection page first. Native format: glyphs as vector outlines, no text layer
(key format 'outline'). Docs 01/02 are the trio with f05/f07 01/02 (_sap_common.TRIO).
"""

import random
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

sys.path[:0] = [str(Path(__file__).resolve().parents[1]), str(Path(__file__).resolve().parent)]
import _sap_common as S  # noqa: E402
import core  # noqa: E402
from core import Cell  # noqa: E402

FAMILY = Path(__file__).stem
FIELDS = ["asset_no", "sub_no", "name", "curr", "acq_date", "ahk_open", "kum_afa_open", "bw_prev", "zugang", "afa_gj",
          "abgang", "afa_abgang", "umbuchung", "bw_close", "ahk_close", "kum_afa_close"]
HEADS = ["Anlage", "UNr.", "Anlagenbezeichnung", "Währg", "Aktivdatum", "AHK GJ-Beg", "AfA GJ-Beg", "Buchwert GJ-Beg",
         "Zugang", "AfA des Jahres", "Abgang", "AfA Abgang", "Umbuchung", "lfd Buchwert", "aktuelle AHK",
         "kumulierte AfA"]
WIDTHS = [24, 17, 92, 20, 38] + [51] * 11
NAME_CHARS = 26  # the print cuts descriptions here, mid-word
MUT_FIELDS = ["asset_no", "sub_no", "name", "acq_date", "ahk_open", "zugang", "abgang", "umbuchung", "ahk_close",
              "kum_afa_open", "afa_gj", "afa_abgang", "kum_afa_close", "bw_prev", "bw_close", "curr"]
MUT_HEADS = ["Anlage", "UNr", "Bezeichnung", "Akt.-Datum", "AHK Beginn", "Zugänge", "Abgänge", "Umbuchungen",
             "AHK Ende", "AfA kum. Beginn", "AfA lfd. Jahr", "AfA auf Abgänge", "AfA kum. Ende", "BW Beginn", "BW Ende",
             "Whrg"]
MUT_WIDTHS = [24, 14, 92, 38] + [51] * 11 + [20]
SUBTITLE = "Gitter HGB §268 Abs.2 - 13 Spalten (breit, unvollständig)"


def table(c: core.Client, mode: str = "print") -> core.Table:
    """mode: 'print' (the list), 'xlsx' (ALV export: group fields as columns, full descriptions), 'mutation'."""
    fields, heads, widths = {"print": (FIELDS, HEADS, WIDTHS), "mutation": (MUT_FIELDS, MUT_HEADS, MUT_WIDTHS),
                             "xlsx": (["bilpos", "konto", "klasse", *FIELDS],
                                      ["Bilanzposition", "Bestandskonto", "Anlagenklasse", *HEADS], [40, 30, 30, *WIDTHS])}[mode]
    t = core.Table(fields=fields, col_widths=widths, align=["r" if f in core.AMOUNTS else "l" for f in fields],
                   header_rows=[[Cell(h) for h in heads]], font="Courier", font_size=5.6, rows_per_page=50)
    S.titles(t, c, "Anlagengitter gebuchte AfA  Bereich 01 Handelsbilanz", SUBTITLE)
    t.neg = "leading" if mode == "mutation" else "trailing"
    rng = random.Random(f"f06-{c.seed}")
    for i, k in enumerate(k for k in S.ledger_classes(c) if any(a.cls == k.code for a in c.assets)):
        A = [a for a in c.assets if a.cls == k.code]
        if i and mode == "print":
            t.add("page_break", {})
        for a in A:
            part = f"{rng.randrange(100, 1000)}-{rng.randrange(100_000, 1_000_000)} "  # internal tool/part number
            full = (part if rng.random() < 0.3 else "") + a.name
            r = t.add_asset(replace(a, asset_no=str(int(a.asset_no) - 1_000_000)))  # SAP: plain running number
            r.cells[fields.index("name")] = r.truth["name"] = full if mode == "xlsx" else full[:NAME_CHARS]
            if mode == "xlsx":
                r.cells[:3] = [S.POS7[k.position], k.account, k.code]
        label = f"Summe Klasse {k.code}" if mode == "mutation" else f"*Klasse {k.code} {k.name}"
        t.add("class_subtotal", {fields.index("asset_no"): label, **S.sums(A, fields)}, {"year": c.year, "class": k.code})
    return S.finish(t)


def look(c: core.Client, mode: str = "print") -> S.Look:
    if mode == "mutation":  # upright landscape, no selection page, no group block
        return S.Look(font="Courier", size=5.6, W=842, H=595, box="grid")
    ks = {k.code: k for k in c.classes}

    def block(r):
        k = ks[r.truth["class"]]
        return (["Buchungskreis", "GeschBereich", "Bilanzposition", "BestandskontoAHK", "Anlagenklasse"],
                [S.bukr(c), "", S.POS7[k.position], k.account, f"{k.code} {k.name}"])

    return S.Look(font="Courier", size=5.6, W=791, H=595, rotate=True, box="grid", block=block, cover=cover(c))


def cover(c: core.Client):
    def tabs(t: core.Table) -> list:
        return [[["Selektionsfeld", "Wert"], ["Buchungskreis", S.bukr(c)], ["Berichtsdatum", core.fmt_date(t.report_date)],
                 ["Bewertungsbereich 1", "01"], ["Sortiervariante", "ZK4"], ["Anlageneinzelliste", "X"],
                 ["ALV Grid Control benutzen", "X"], ["Gitterversion", "ZG13"], ["Gebuchte Abschreibung", "X"]],
                [["Sortierkriterien", "Aufst.", "Abst.", "Zwischensumme"], ["Bilanzposition", "X", "", ""],
                 ["Bestandskonto AHK", "X", "", ""], ["Anlagenklasse", "X", "", "X"]],
                [["Datenstatistik", "Anzahl"], ["Übergebene Sätze", str(sum(r.kind == "asset" for r in t.rows))],
                 ["Ermittelte Summenzeilen", str(sum(r.kind == "class_subtotal" for r in t.rows))]]]
    return tabs


def build(out: Path) -> None:
    d = out / FAMILY
    d.mkdir(parents=True, exist_ok=True)

    def key(n, c, fmt, variant, t, planted=(), **kw):
        S.write_key(d, FAMILY, n, c, fmt, variant, t, planted, row_level="asset", **kw)

    def via_print(t, lk, then):  # render to a temp print, then outline / scan it
        with tempfile.TemporaryDirectory() as tmp:
            S.render(t, Path(tmp) / "print.pdf", lk)
            then(Path(tmp) / "print.pdf")

    c = S.client(S.TRIO_SEED, year=2025, n_classes=13, assets_per_class=(4, 9))  # 01 + 02: the trio ledger
    t = table(c)
    via_print(t, look(c), lambda p: core.outline(p, d / f"{FAMILY}_01.pdf"))
    key(1, c, "outline", "perfect", t, trio=True)
    t = table(c, "xlsx")
    core.render_xlsx(t, d / f"{FAMILY}_02.xlsx", sheet="Anlagengitter", number_format="#,##0.00;#,##0.00-",
                     extra_sheets={"Selektion": [row for tab in cover(c)(t) for row in [*tab, []]]})
    key(2, c, "xlsx", "perfect", t, trio=True)

    c = S.client(522, year=2025, n_classes=8, assets_per_class=(3, 8))  # 03 scan, two pages fed in turned
    t = table(c)
    via_print(t, look(c), lambda p: core.scan(p, d / f"{FAMILY}_03.pdf", dpi=200, rotate={2: 180, 4: 90}, seed=603))
    key(3, c, "scan", "perfect", t)

    c, rng = S.client(523, year=2025, n_classes=9, assets_per_class=(5, 12)), random.Random(523)  # 04 text layer
    t = table(c)
    P = [core.p_glued(t, rng)[1]]
    S.finish(t)
    P.append(core.p_non_tying(t, rng)[1])
    S.finish(t)
    P.append(core.p_carry_over(t)[1])
    S.finish(t)
    S.render(t, d / f"{FAMILY}_04.pdf", look(c))
    key(4, c, "pdf", "problem", t, P)

    c = S.client(524, year=2024, fy_end=(3, 31), n_classes=10, assets_per_class=(3, 7))  # 05 outline, header once
    t = table(c)
    P = [core.p_no_repeat_header(t)[1], core.p_noise(t, c, "unrelated", "middle")[1]]
    via_print(t, look(c), lambda p: core.outline(p, d / f"{FAMILY}_05.pdf"))
    key(5, c, "outline", "problem", t, P)

    c = S.client(525, year=2025, n_classes=7, assets_per_class=(4, 10))  # 06 last column cut at the page edge
    t = table(c)
    for r in t.rows:
        if r.truth:
            r.flags["kum_afa_close"] = "cut_off"
    S.render(t, d / f"{FAMILY}_06.pdf", replace(look(c), cut=4))
    key(6, c, "pdf", "problem", t, [{
        "type": "column_cut_off", "where": [{"fields": ["kum_afa_close"], "pages": "all", "chars_lost": 4}],
        "expected_behaviour": "the right-most column (kumulierte AfA) runs past the page edge and loses its last "
                              "characters (cents, sign); rebuild it: kumulierte AfA = AfA GJ-Beg + AfA des Jahres + "
                              "AfA Abgang = lfd Buchwert - aktuelle AHK (as printed, with signs)"}])

    c = S.client(526, year=2025, n_classes=11, assets_per_class=(3, 7))  # 07 reordered, renamed, leading minus
    t = table(c, "mutation")
    S.render(t, d / f"{FAMILY}_07.pdf", look(c, "mutation"))
    key(7, c, "pdf", "layout-mutation", t)
