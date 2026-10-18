"""Layout fuzzer: documents in layouts the pipeline was never tuned on, each with an exact answer key.

Built on the corpus generator (A3-Hackathon-Corpus/_generator/core.py): the same invented clients and accounting
laws, but the layout is drawn at random: which columns, in which order, how headers are worded (German variants,
abbreviations, English), group headers or not, number and zero styles, font, orientation, how totals are labelled and
where they sit, which problems are planted, and the format (xlsx, text PDF, scan, glyph outlines).

    .venv/bin/python bench/fuzz.py bench/dev  --seeds 1000-1059    # develop against these; read the failures
    .venv/bin/python bench/fuzz.py bench/test --seeds 5000-5079    # held out: only ever look at the score
"""

from __future__ import annotations

import argparse
import random
import sys
import tempfile
import traceback
from datetime import date
from pathlib import Path

GEN = Path(__file__).resolve().parents[2] / "A3-Hackathon-Corpus" / "_generator"
sys.path.insert(0, str(GEN))
import core  # noqa: E402
from core import Cell  # noqa: E402

GROUP = {"AHK": ["Anschaffungs- und Herstellungskosten", "Anschaffungs-/Herstellungskosten", "AHK", "Bruttowerte",
                 "Historische Anschaffungskosten", "Acquisition and production cost"],
         "AFA": ["Kumulierte Abschreibungen", "Abschreibungen", "Wertberichtigungen", "kum. AfA", "Accumulated depreciation"],
         "BW": ["Buchwerte", "Restbuchwerte", "Nettobuchwerte", "Net book value"]}
ROLE = {"open": ["{d0}", "Stand {d0}", "Anfangsstand", "Beginn GJ", "Vortrag", "Opening"],
        "add": ["Zugänge", "Zugang", "Additions"], "disposal": ["Abgänge", "Abgang", "Disposals"],
        "transfer": ["Umbuchungen", "Umbuchung", "Umgliederungen", "Transfers"],
        "writeup": ["Zuschreibungen", "Zuschreibung", "Wertaufholung", "Write-ups"],
        "year": ["Geschäftsjahr", "lfd. Jahr", "des Jahres", "Zuführung", "Charge for the year"],
        "close": ["{d1}", "Stand {d1}", "Endstand", "Ende GJ", "Closing"], "prev": ["{dp}", "Vorjahr", "Stand {dp}", "Prior year"]}
FULL = {"ahk_open": ["AHK {d0}", "Anschaffungskosten {d0}", "AK/HK Beginn", "AHK GJ-Beginn", "Bruttowert Anfang", "Cost {d0}"],
        "zugang": ["Zugänge", "Zugang GJ", "Zugänge Geschäftsjahr", "Additions", "Investitionen"],
        "abgang": ["Abgänge", "Abgang AHK", "Abgänge zu AK", "Disposals at cost"],
        "abgang_bw": ["Abgänge zum Restbuchwert", "Abgang RBW", "Disposals (NBV)"],
        "umbuchung": ["Umbuchungen", "Umbuchung", "Umgliederungen", "Transfers"],
        "zuschreibung": ["Zuschreibungen", "Zuschreibung", "Wertaufholungen", "Write-ups"],
        "ahk_close": ["AHK {d1}", "Anschaffungskosten {d1}", "AK/HK Ende", "AHK GJ-Ende", "Bruttowert Ende", "Cost {d1}"],
        "kum_afa_open": ["kum. AfA {d0}", "Abschreibungen kumuliert {d0}", "WB Beginn", "Acc. depreciation {d0}"],
        "afa_gj": ["AfA GJ", "Abschreibungen Geschäftsjahr", "AfA lfd. Jahr", "planm. AfA", "Depreciation for the year"],
        "afa_abgang": ["AfA auf Abgänge", "Abgang kum. AfA", "AfA Abgang", "Depreciation on disposals"],
        "kum_afa_close": ["kum. AfA {d1}", "Abschreibungen kumuliert {d1}", "WB Ende", "Acc. depreciation {d1}"],
        "bw_close": ["Buchwert {d1}", "Restbuchwert {d1}", "RBW Ende", "Net book value {d1}", "Buchwert lfd. Jahr"],
        "bw_prev": ["Buchwert {dp}", "Restbuchwert Vorjahr", "RBW Vorjahr", "Net book value {dp}", "Buchwert VJ"]}
FIELD_ROLE = {"ahk_open": ("AHK", "open"), "zugang": ("AHK", "add"), "abgang": ("AHK", "disposal"),
              "umbuchung": ("AHK", "transfer"), "ahk_close": ("AHK", "close"), "kum_afa_open": ("AFA", "open"),
              "afa_gj": ("AFA", "year"), "afa_abgang": ("AFA", "disposal"), "zuschreibung": ("AFA", "writeup"),
              "kum_afa_close": ("AFA", "close"), "bw_close": ("BW", "close"), "bw_prev": ("BW", "prev")}
ID_LABEL = {"asset_no": ["Inv.-Nr.", "Anlage", "Nr.", "Anlagen-Nr.", "Asset no."],
            "name": ["Bezeichnung", "Anlagenbezeichnung", "Gegenstand", "Description"],
            "acq_date": ["Zugangsdatum", "Aktivierung", "AK-Datum", "Acquired"]}
SUM_LABEL = ["Summe {name}", "Summe Konto {account}", "Total {name}", "{name} gesamt", "*", ""]
GRAND_LABEL = ["Summe Anlagevermögen", "Gesamt", "Total", "Anlagevermögen gesamt", "Gesamtsumme", ""]
HEADING = ["{account} {name}", "Konto {account}", "Klasse {code} {name}", "{name}", None]


def _schema(rng: random.Random) -> tuple[str, list[str]]:
    kind = rng.choices(["gross", "net", "register"], [0.65, 0.2, 0.15])[0]
    if kind == "gross":
        f = ["ahk_open", "zugang", "abgang", "umbuchung", "ahk_close", "kum_afa_open", "afa_gj", "afa_abgang",
             "zuschreibung", "kum_afa_close", "bw_close", "bw_prev"]
        for opt in ("umbuchung", "zuschreibung", "ahk_close", "kum_afa_open", "afa_abgang", "bw_prev"):
            if rng.random() < 0.25:
                f.remove(opt)
        if "zuschreibung" in f and rng.random() < 0.3:  # some layouts print write-ups among the cost movements
            f.remove("zuschreibung")
            f.insert(f.index("ahk_close") if "ahk_close" in f else 4, "zuschreibung")
    elif kind == "net":
        f = ["bw_prev", "zugang", "abgang_bw", "umbuchung", "zuschreibung", "afa_gj", "kum_afa_close", "bw_close"]
        for opt in ("umbuchung", "zuschreibung", "kum_afa_close"):
            if rng.random() < 0.3:
                f.remove(opt)
    else:
        f = rng.choice([["ahk_close", "bw_close"], ["ahk_close", "kum_afa_close", "bw_close"],
                        ["ahk_open", "zugang", "abgang", "ahk_close", "bw_close"]])
    if rng.random() < 0.2 and "bw_prev" in f and "bw_close" in f:  # book values printed prior year first
        f.remove("bw_prev")
        f.insert(f.index("bw_close"), "bw_prev")
    if rng.random() < 0.2:  # one adjacent swap inside the movements
        movers = [i for i in range(len(f) - 1) if {f[i], f[i + 1]} <= {"zugang", "abgang", "umbuchung", "zuschreibung",
                                                                      "afa_gj", "afa_abgang", "abgang_bw"}]
        if movers:
            i = rng.choice(movers)
            f[i], f[i + 1] = f[i + 1], f[i]
    return kind, f


def _label(text: str, c: core.Client, style: str) -> str:
    d0, d1 = c.fy_start, c.fy_end
    dp = d1.replace(year=d1.year - 1)
    return text.format(d0=core.fmt_date(d0, style), d1=core.fmt_date(d1, style), dp=core.fmt_date(dp, style))


def build_table(c: core.Client, rng: random.Random, spec: dict) -> core.Table:
    amount_fields = spec["fields"]
    ids = ["asset_no", "name"] + (["acq_date"] if spec["acq_date"] else [])
    fields = ids + amount_fields if not spec["date_last"] else ["asset_no", "name"] + amount_fields + ["acq_date"]
    W = 790 if spec["landscape"] else 545
    id_w = {"asset_no": 34, "name": 130, "acq_date": 44}
    aw = max(34, min(60, (W - sum(id_w[i] for i in ids)) / len(amount_fields)))
    widths = [id_w.get(f, aw) for f in fields]
    align = ["l" if f in ("asset_no", "name") else ("c" if f == "acq_date" else "r") for f in fields]
    ds = spec["date_style"]
    labels = {}
    for f in fields:
        if f in ID_LABEL:
            labels[f] = rng.choice(ID_LABEL[f])
        elif spec["groups"] and f in FIELD_ROLE:
            labels[f] = _label(rng.choice(ROLE[FIELD_ROLE[f][1]]), c, ds)
        else:
            labels[f] = _label(rng.choice(FULL[f]), c, ds)
    header_rows = []
    if spec["groups"]:  # group header over contiguous blocks
        row, i = [], 0
        while i < len(fields):
            g = FIELD_ROLE.get(fields[i], (None,))[0]
            j = i
            while j < len(fields) and FIELD_ROLE.get(fields[j], (None,))[0] == g:
                j += 1
            row.append(Cell(rng.choice(GROUP[g]) if g else "", j - i))
            i = j
        header_rows.append(row)
    two = [labels[f].split(" ", 1) if (len(labels[f]) * spec["font_size"] * 0.5 > widths[k] and " " in labels[f]) else [labels[f]]
           for k, f in enumerate(fields)]
    if any(len(x) == 2 for x in two):
        header_rows.append([Cell(x[0]) for x in two])
        header_rows.append([Cell(x[1] if len(x) == 2 else "") for x in two])
    else:
        header_rows.append([Cell(labels[f]) for f in fields])
    if spec["unit_row"]:
        header_rows.append([Cell("EUR" if f not in ids else "") for f in fields])
    doc_title = rng.choice(["Anlagenspiegel", "Anlagengitter", "Entwicklung des Anlagevermögens", "Anlagenverzeichnis",
                            "Fixed asset schedule"])
    period = rng.choice([f"zum {core.fmt_date(c.fy_end)}", f"vom {core.fmt_date(c.fy_start)} bis {core.fmt_date(c.fy_end)}",
                         f"Geschäftsjahr {core.fmt_date(c.fy_start)} - {core.fmt_date(c.fy_end)}", f"per {core.fmt_date(c.fy_end)}"])
    t = core.Table(fields=fields, col_widths=widths, align=align, header_rows=header_rows,
                   title_lines=[c.company, f"{doc_title} {period}"] + (["Beträge in EUR"] if rng.random() < 0.5 else []),
                   num_style=spec["num_style"], zero=spec["zero"], date_style=ds, font=spec["font"],
                   font_size=spec["font_size"], landscape=spec["landscape"], repeat_header=True,
                   page_header=rng.choice(["", "Seite {p} von {n}", "Blatt {p}"]),
                   page_footer=rng.choice(["Seite {p} von {n}", "Druckdatum {date}", ""]),
                   print_date=date(c.year + 1, rng.randint(1, 5), rng.randint(1, 28)), rules=spec["rules"])
    for k, assets in c.by_class():
        names = {"name": k.name, "account": k.account, "code": k.code}
        if spec["sum_place"] == "above":
            t.rows.append(t.total("class_subtotal", spec["sum_label"].format(**names) or k.name, assets, cls=k.code))
        elif spec["heading"]:
            t.add("heading", {"asset_no" if "asset_no" in fields else 0: "", "name": spec["heading"].format(**names)})
        for a in assets:
            t.add_asset(a)
        if spec["sum_place"] == "below":
            t.add_total("class_subtotal", spec["sum_label"].format(**names), assets, cls=k.code)
    if spec["grand"] is not None:
        if spec["sum_place"] == "above":
            t.rows.insert(0, t.total("grand_total", spec["grand"] or "Summe", c.assets))
        else:
            t.add_total("grand_total", spec["grand"], c.assets)
    return t


def build(out: Path, seed: int) -> str:
    rng = random.Random(f"fuzz-{seed}")
    c = core.make_client(seed, year=rng.choice([2024, 2025]), fy_end=rng.choice([(12, 31)] * 4 + [(6, 30), (9, 30)]),
                         n_classes=rng.randint(3, 10), assets_per_class=(2, rng.randint(4, 9)),
                         subnumbers=rng.random() < 0.2, disposals_at="cost")
    kind, fields = _schema(rng)
    fmt = rng.choices(["xlsx", "pdf", "scan", "outline"], [0.25, 0.4, 0.2, 0.15])[0]
    spec = {"fields": fields, "acq_date": rng.random() < 0.6, "date_last": rng.random() < 0.15,
            "groups": rng.random() < 0.4 and kind == "gross", "unit_row": rng.random() < 0.4,
            "num_style": rng.choices(["de", "de_nodots", "de_trailing_minus", "paren", "en"], [0.6, 0.1, 0.12, 0.1, 0.08])[0],
            "zero": rng.choice([None, "", "-", None]), "date_style": rng.choice(["de", "de", "de_short"]),
            "font": rng.choice(["Helvetica", "Courier", "Times-Roman", "Helvetica"]), "font_size": rng.choice([6.5, 7, 7.5, 8]),
            "landscape": len(fields) > 6 or rng.random() < 0.5, "rules": rng.random() < 0.5,
            "sum_place": rng.choices(["below", "above", "none"], [0.75, 0.1, 0.15])[0],
            "sum_label": rng.choice(SUM_LABEL), "heading": rng.choice(HEADING),
            "grand": rng.choice(GRAND_LABEL) if rng.random() < 0.7 else None}
    multi = rng.random() < 0.12
    clients = c.years(2) if multi else [c]
    tables = [build_table(cc, random.Random(f"fuzz-{seed}-{cc.year}"), spec) for cc in clients]
    planted = []
    if rng.random() < 0.06:
        tables = [core.p_teur(t)[0] for t in tables]
        planted.append({"type": "teur_rounding"})
    t = tables[-1]
    for name, p, fn in (("glued", 0.15, lambda: core.p_glued(t, rng)), ("non_tying", 0.2, lambda: core.p_non_tying(t, rng)),
                        ("wrap", 0.15, lambda: core.p_wrap(t, rng))):
        if rng.random() < p and fmt != "xlsx":
            try:
                planted.append(fn()[1])
            except Exception:
                pass
    if fmt != "xlsx":
        if rng.random() < 0.12:
            try:
                planted.append(core.p_no_repeat_header(t)[1])
            except Exception:
                pass
        if rng.random() < 0.1:
            try:
                planted.append(core.p_carry_over(t)[1])
            except Exception:
                pass
        if rng.random() < 0.15:
            try:
                planted.append(core.p_noise(t, c, rng.choice(["cover", "blank", "unrelated"]))[1])
            except Exception:
                pass
    if multi:
        planted.append(core.p_multi_year(tables)[1])
    doc_id = f"fz_{seed}"
    path = out / f"{doc_id}.{'xlsx' if fmt == 'xlsx' else 'pdf'}"
    if fmt == "xlsx":
        core.render_xlsx(tables, path, sheet=[f"AV {cc.year}" for cc in clients] if multi and rng.random() < 0.5 else "Anlagen",
                         numbers=rng.choice(["number"] * 4 + ["text"]))
    elif fmt == "pdf":
        core.render_pdf(tables, path)
    else:
        with tempfile.TemporaryDirectory() as tmp:
            core.render_pdf(tables, Path(tmp) / "print.pdf")
            (core.scan(Path(tmp) / "print.pdf", path, seed=seed) if fmt == "scan" else core.outline(Path(tmp) / "print.pdf", path))
    meta = core.key_meta(c, doc_id=doc_id, family="fuzz", doc_type="Anlagenspiegel", format=fmt,
                         variant="problem" if planted else "layout-mutation", schema=kind, spec={k: v for k, v in spec.items()})
    if multi:
        meta["years"] = [cc.year for cc in clients]
    core.write_key(out / f"{doc_id}.key.json", meta, tables, planted)
    return f"{doc_id}: {fmt:7} {kind:8} {len(fields):2} cols groups={spec['groups']} num={spec['num_style']} sums={spec['sum_place']} planted={[p['type'] for p in planted if p]}"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("out", type=Path)
    ap.add_argument("--seeds", default="1000-1029")
    a = ap.parse_args()
    lo, hi = map(int, a.seeds.split("-"))
    a.out.mkdir(parents=True, exist_ok=True)
    for s in range(lo, hi + 1):
        try:
            print(build(a.out, s), flush=True)
        except Exception as e:
            print(f"fz_{s}: FAILED {type(e).__name__}: {e}")
            traceback.print_exc(limit=2)


if __name__ == "__main__":
    main()
