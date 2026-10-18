"""Fill the Arbeitsversion template straight from the corpus answer keys (<doc_id>.key.json).

This is the target output for each document: what a perfect extraction of the schedule writes into the
Arbeitsversion. Every amount is the key's value (int cents -> EUR); its cell comment names the source position and
the printed text. Key flags, planted irregularities and multi-year content become lines on 'Offene Punkte'.

    .venv/bin/python key_to_arbeitsversion.py ../A3-Hackathon-Corpus out_key
    .venv/bin/python key_to_arbeitsversion.py <corpus> <out> --template path/to/Arbeitsversion.xlsx --only f02_ f09_

Writes <out>/Arbeitsversion_<doc_id>.xlsx per key and <out>/summary.json.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment

TEMPLATE = Path(__file__).resolve().parent.parent / "Arbeitsversion-merged-layout-20.xlsx"

AV_COLS = {"ahk_open": "W", "zugang": "X", "abgang": "Y", "umbuchung": "Z", "ahk_close": "AA", "kum_afa_open": "AB",
           "afa_gj": "AC", "afa_abgang": "AE", "zuschreibung": "AF", "kum_afa_close": "AH", "bw_close": "AI",
           "bw_prev": "AJ"}
GT_COLS = {"ahk_open": "D", "zugang": "E", "abgang": "F", "umbuchung": "G", "ahk_close": "I", "kum_afa_open": "J",
           "afa_gj": "L", "afa_abgang": "N", "zuschreibung": "O", "kum_afa_close": "R", "bw_close": "S", "bw_prev": "T"}
GT_NET_COLS = {"bew_ahk": "H", "bew_wb": "Q"}
GT_FIRST, GT_ROWS, GT_SUM_ROW = 8, 20, 28
GT_FORMULAS = {"I": ("D", "E", "F", "G", "={D}+{E}-{F}+{G}"), "R": ("J", "L", "N", "O", "={J}+{L}-{N}-{O}"),
               "S": ("I", "R", "={I}-{R}"), "T": ("D", "J", "={D}-{J}")}
METHOD = {"linear": "linear", "degressiv": "degressiv", "GWG": "GWG", "keine": "keine AfA"}
HGB = {"A.I": "Immaterielle Vermögensgegenstände",
       "A.I.1": "Selbst geschaffene gewerbliche Schutzrechte und ähnliche Rechte und Werte",
       "A.I.2": "Entgeltlich erworbene Konzessionen, gewerbliche Schutzrechte und ähnliche Rechte und Werte sowie Lizenzen",
       "A.I.3": "Geschäfts- oder Firmenwert", "A.I.4": "Geleistete Anzahlungen",
       "A.II": "Sachanlagen", "A.II.1": "Grundstücke, grundstücksgleiche Rechte und Bauten",
       "A.II.2": "Technische Anlagen und Maschinen", "A.II.3": "Andere Anlagen, Betriebs- und Geschäftsausstattung",
       "A.II.4": "Geleistete Anzahlungen und Anlagen im Bau",
       "A.III": "Finanzanlagen", "A.III.1": "Anteile an verbundenen Unternehmen",
       "A.III.2": "Ausleihungen an verbundene Unternehmen", "A.III.3": "Beteiligungen",
       "A.III.4": "Ausleihungen an Unternehmen, mit denen ein Beteiligungsverhältnis besteht",
       "A.III.5": "Wertpapiere des Anlagevermögens", "A.III.6": "Sonstige Ausleihungen"}
FLAG_TEXT = {"glued": "im Original mit dem Nachbarbetrag zusammengeschrieben",
             "cut_off": "im Original am Seitenrand abgeschnitten",
             "non_tying": "Summe des Mandanten stimmt nicht mit ihren Zeilen überein (wie gedruckt übernommen)",
             "transfer_incomplete": "Umbuchungen gleichen sich nicht aus; Wert laut Schlüssel statt gedruckt"}


def is_code(s: str | None) -> bool:
    return bool(s) and re.fullmatch(r"[A-H]\.[IVX]+(\.\d+)?", s) is not None


def eur(cents: int | None) -> float | None:
    return None if cents is None else cents / 100


def source_name(meta: dict) -> str:
    return f"{meta['doc_id']}.{'xlsx' if meta['format'] == 'xlsx' else 'pdf'}"


def where(ptr: dict, field: str | None = None) -> str:
    if ptr.get("sheet"):
        return f"{ptr['sheet']}!{(ptr.get('cells') or {}).get(field) or ptr.get('cell')}"
    return f"Seite {ptr.get('page')}, Zeile {ptr.get('line')}"


def note(ws, ref: str, src: str, r: dict, field: str) -> None:
    printed = (r.get("printed") or {}).get(field)
    text = f"Quelle: {src}, {where(r['pointer'], field)}" + (f" · gedruckt '{printed}'" if printed else " · leer gedruckt (= 0)")
    flag = (r.get("flags") or {}).get(field)
    if flag:
        text += f" · {FLAG_TEXT.get(flag, flag)}"
    c = Comment(text, "Datenübernahme")
    c.width, c.height = 320, 70
    ws[ref].comment = c


def identity(r: dict) -> tuple:
    return (r["kind"], r.get("level"), r.get("class"), r.get("account"), r.get("position"), r.get("name"), r.get("asset_no"),
            r.get("sub_no"))


def dedupe(rows: list[dict], open_points: list, src: str) -> list[dict]:
    """Summary blocks printed twice (cover summary + detail, summary sheet) appear once; differing repeats are noted."""
    first: dict[tuple, dict] = {}
    out = []
    for r in rows:
        k = identity(r)
        if k in first:
            if first[k]["fields"] != r["fields"]:
                open_points.append(("—", where(r["pointer"]), f"Zeile '{label(r)}' wird mehrfach mit abweichenden "
                                                               f"Beträgen gedruckt; übernommen ist die erste ({where(first[k]['pointer'])})."))
            continue
        first[k] = r
        out.append(r)
    return out


def label(r: dict) -> str:
    if r.get("name"):
        return r["name"]
    if r["kind"] == "grand_total":
        return "Gesamtsumme"
    if r["kind"] == "position_subtotal":
        p = r.get("position") or r.get("class")
        return HGB.get(p, p) if is_code(p) else p
    if r["kind"] == "class_subtotal":
        return r["class"] if not (r.get("class") or "").isdigit() else f"Anlagenklasse {r['class']}"
    if r["kind"] == "account_subtotal":
        return f"Konto {r['account']}"
    return r.get("class") or r.get("account") or r["kind"]


def write(template: Path, out: Path, key: dict) -> dict:
    meta, src = key["meta"], source_name(key["meta"])
    wb = openpyxl.load_workbook(template)
    open_points: list[tuple[str, str, str]] = []

    years = sorted({r["year"] for r in key["rows"]})
    year = years[-1]
    period_end = date(year, int(meta["period_end"][5:7]), int(meta["period_end"][8:10]))
    if len(years) > 1:
        open_points.append(("—", src, f"Die Übergabe enthält die Jahre {', '.join(map(str, years))}; übernommen ist {year}, "
                                     f"die übrigen stehen nur im Schlüssel."))
    rows = [r for r in key["rows"] if r["year"] == year]
    if any(r["kind"] == "carry_over" for r in rows):
        open_points.append(("—", src, "Überträge (Übertrag/Vortrag) zwischen den Seiten sind keine Zeilen des Gitters und "
                                     "nicht übernommen."))
        rows = [r for r in rows if r["kind"] != "carry_over"]
    rows = dedupe(rows, open_points, src)
    net_abgang = any("abgang_bw" in r["fields"] and "abgang" not in r["fields"] for r in rows)

    # the class each asset sits under -> its position, for 'Bilanzposition laut Mandant' on the class rows
    pos_of_class = {}
    codes: dict[tuple, set] = {}  # ('class', x) / ('position', x) -> § 266 codes of the assets beneath it
    for r in rows:
        if r["kind"] == "asset" and r.get("class") and r.get("position"):
            pos_of_class.setdefault(r["class"], Counter())[r["position"]] += 1
        code = r.get("hgb_position") or (r.get("position") if is_code(r.get("position")) else None)
        if r["kind"] == "asset" and code:
            for k in ("class", "position", "account"):
                if r.get(k):
                    codes.setdefault((k, r[k]), set()).add(code)
    pos_of_class["_codes"] = codes

    # ---------------------------------------------------------------- Anlagenverzeichnis
    av = wb["Anlagenverzeichnis"]
    assets = [r for r in rows if r["kind"] == "asset" and r.get("asset_no")]
    for n, r in enumerate(assets, start=1):
        i = n + 1
        sub = r.get("sub_no") or ""
        av[f"A{i}"] = r["asset_no"] + (f"/{sub}" if sub and sub != "0" else "")
        av[f"B{i}"] = r.get("name")
        cls = r.get("class")
        if cls:
            av[f"D{i}" if cls.isdigit() else f"C{i}"] = cls
        pos = r.get("hgb_position") or (r.get("position") if is_code(r.get("position")) else None)
        if pos:
            av[f"E{i}"] = pos
        if r.get("position") and not is_code(r["position"]):
            av[f"G{i}"] = r["position"]
        if r.get("account"):
            av[f"H{i}"] = r["account"]
        if r.get("acq_date"):
            av[f"K{i}"] = date.fromisoformat(r["acq_date"])
            av[f"K{i}"].number_format = "DD.MM.YYYY"
        if r.get("nd_months") is not None:
            av[f"O{i}"] = r["nd_months"]
            av[f"N{i}"] = round(r["nd_months"] / 12, 2)
        if r.get("method"):
            av[f"P{i}"] = METHOD.get(r["method"], r["method"])
            av[f"AL{i}"] = "Ja" if r["method"] == "GWG" else "Nein"
        for f, col in AV_COLS.items():
            if f in r["fields"]:
                av[f"{col}{i}"] = eur(r["fields"][f])
                note(av, f"{col}{i}", src, r, f)
        av[f"AQ{i}"] = f"{meta['doc_id']}#{key['rows'].index(r)}"
        for f, flag in (r.get("flags") or {}).items():
            col = AV_COLS.get(f, "A")
            open_points.append(("Anlagenverzeichnis", f"{col}{i}", f"Anlage {av[f'A{i}'].value}, {f}: {FLAG_TEXT.get(flag, flag)}."))
    if assets and net_abgang:
        open_points.append(("Anlagenverzeichnis", "Y", "Abgang nur zum Restbuchwert gedruckt (Nettodarstellung): "
                                                     "Bruttodarstellung nach § 284 Abs. 3 HGB nachfordern."))

    # ---------------------------------------------------------------- Gitter
    by_bilanzposition = meta.get("row_level") == "bilanzposition"
    grand = [r for r in rows if r["kind"] == "grand_total"]
    if by_bilanzposition:
        klasse_rows: list[tuple[dict, str]] = []
        position_rows = [(r, "Bilanzposition" if r["kind"] == "asset" else "Zwischensumme Bilanzposition")
                         for r in rows if r["kind"] in ("asset", "position_subtotal")]
    else:
        if any(r["kind"] == "class_subtotal" for r in rows):
            is_class = lambda r: r["kind"] == "class_subtotal"
        elif any(r["kind"] == "asset" and not r.get("asset_no") for r in rows):
            is_class = lambda r: r["kind"] == "asset" and not r.get("asset_no")  # SAP class prints, net Spiegel
        else:
            is_class = lambda r: r["kind"] == "account_subtotal"  # only account subtotals printed
        klasse_rows = [(r, "Anlagenklasse" if is_class(r) else "Zwischensumme Bilanzposition")
                       for r in rows if is_class(r) or r["kind"] == "position_subtotal"]
        position_rows = [(r, "Bilanzposition") for r in rows if r["kind"] == "position_subtotal"]

    gitter_info = {}
    for gname, body in (("Gitter nach Anlagenklasse", klasse_rows), ("Gitter nach Bilanzposition", position_rows)):
        gt = wb[gname]
        gt["C2"] = period_end
        has_total = grand and gname == ("Gitter nach Bilanzposition" if by_bilanzposition else "Gitter nach Anlagenklasse")
        if not body and not has_total:
            gitter_info[gname] = 0
            continue
        if len(body) > GT_ROWS:
            gt.insert_rows(GT_FIRST + GT_ROWS, amount=len(body) - GT_ROWS)
        sum_row = GT_SUM_ROW + max(0, len(body) - GT_ROWS)
        for k, (r, zeilenart) in enumerate(body):
            fill_gitter(gt, gname, GT_FIRST + k, r, zeilenart, src, meta, key, pos_of_class, open_points)
        if grand:
            g = grand[-1]  # the outermost total ('****' after '***'); equal repeats were dropped above
            fill_gitter(gt, gname, sum_row, g, "Gesamtsumme", src, meta, key, pos_of_class, open_points)
            gt[f"C{sum_row}"] = "Summe laut Mandant"
            for other in grand[:-1]:
                if other["fields"] != g["fields"]:
                    open_points.append((gname, f"C{sum_row}", f"Gesamtsumme {where(other['pointer'])} weicht von "
                                                              f"{where(g['pointer'])} ab; übernommen ist die letzte."))
        if net_abgang:
            open_points.append((gname, "F", "Abgang nur zum Restbuchwert gedruckt (Nettodarstellung): "
                                            "Bruttodarstellung nach § 284 Abs. 3 HGB nachfordern."))
        gitter_info[gname] = len(body)
    if not position_rows:
        open_points.append(("Gitter nach Bilanzposition", "—", "Der Mandant druckt keine Summen je Bilanzposition; "
                                                              "das Blatt bleibt leer."))

    # ---------------------------------------------------------------- Offene Punkte
    for p in key.get("planted", []):
        n = len(p.get("where", [])) if isinstance(p.get("where"), list) else 1
        open_points.append(("—", src, f"Erkannt: {p['type']} ({n} Stelle{'n' if n != 1 else ''})"))
    op = wb["Offene Punkte"]
    if open_points:
        op["A4"] = None
        for k, (tab, ref, question) in enumerate(open_points):
            row = 4 + k
            op[f"A{row}"], op[f"B{row}"], op[f"C{row}"], op[f"D{row}"] = tab, ref, question, "offen (Datenübernahme)"

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"doc_id": meta["doc_id"], "file": out.name, "year": year, "anlagenverzeichnis_rows": len(assets),
            **{k.replace("Gitter nach ", "gitter_").lower(): v for k, v in gitter_info.items()},
            "grand_total": bool(grand), "open_points": len(open_points)}


def derived_code(r: dict, codes: dict[tuple, set]) -> str | None:
    """§ 266 code of a subtotal from the assets beneath it: their one code, else their common parent ('A.II')."""
    k = {"position_subtotal": "position", "account_subtotal": "account"}.get(r["kind"], "class")
    found = codes.get((k, r.get(k))) or (codes.get(("account", r.get("account"))) if r.get("account") else None)
    if not found:
        return None
    if len(found) == 1:
        return next(iter(found))
    parents = {c.rsplit(".", 1)[0] if c.count(".") == 2 else c for c in found}
    return parents.pop() if len(parents) == 1 else None


def fill_gitter(gt, gname, i, r, zeilenart, src, meta, key, pos_of_class, open_points) -> None:
    pos = r.get("hgb_position") or r.get("position") or (r.get("class") if r["kind"] == "position_subtotal" else None)
    if r["kind"] != "grand_total":
        if not is_code(pos):
            pos = derived_code(r, pos_of_class["_codes"])
        if pos:
            gt[f"A{i}"] = pos
        if gname == "Gitter nach Anlagenklasse" and zeilenart == "Anlagenklasse":
            nr = r.get("class") if (r.get("class") or "").isdigit() else r.get("account")
            if nr:
                gt[f"B{i}"] = nr
            mandant_pos = r.get("position") or (pos_of_class.get(r.get("class")) or Counter()).most_common(1)
            if isinstance(mandant_pos, list):
                mandant_pos = mandant_pos[0][0] if mandant_pos else None
            if mandant_pos:
                gt[f"W{i}"] = mandant_pos
                if not gt[f"A{i}"].value and is_code(mandant_pos):
                    gt[f"A{i}"] = mandant_pos
        gt[f"C{i}"] = label(r)
    gt[f"V{i}"] = zeilenart
    gt[f"X{i}"] = f"{meta['doc_id']}#{key['rows'].index(r)}"
    for f, col in GT_COLS.items():
        if f in r["fields"]:
            gt[f"{col}{i}"] = eur(r["fields"][f])
            note(gt, f"{col}{i}", src, r, f)
    for f, col in GT_NET_COLS.items():
        if f in (r.get("net_fields") or {}):
            v = r["net_fields"][f]
            gt[f"{col}{i}"] = eur(-v if f == "bew_wb" else v)  # write-backs and disposals reduce the cumulative AfA
            note(gt, f"{col}{i}", src, r, f)
    for col, (*needs, formula) in GT_FORMULAS.items():
        if gt[f"{col}{i}"].value is None and all(gt[f"{c}{i}"].value is not None for c in needs):
            gt[f"{col}{i}"] = formula.format(**{c: f"{c}{i}" for c in needs})  # 'Mandant, sonst Formel'
    for f, flag in (r.get("flags") or {}).items():
        ref = f"{GT_COLS.get(f, 'C')}{i}"
        open_points.append((gname, ref, f"Zeile '{gt[f'C{i}'].value}', {'alle Felder' if f == '*' else f}: "
                                        f"{FLAG_TEXT.get(flag, flag)}."))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("corpus", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--template", type=Path, default=TEMPLATE)
    ap.add_argument("--only", nargs="*", default=[], help="doc_id prefixes")
    a = ap.parse_args()
    keys = sorted(p for p in a.corpus.rglob("*.key.json")
                  if not any(part.startswith(("_", ".")) for part in p.relative_to(a.corpus).parts))
    if a.only:
        keys = [p for p in keys if p.name.startswith(tuple(a.only))]
    summary = []
    for p in keys:
        key = json.loads(p.read_text(encoding="utf-8"))
        info = write(a.template, a.out / f"Arbeitsversion_{key['meta']['doc_id']}.xlsx", key)
        summary.append(info)
        print(f"{info['doc_id']:28} AV {info['anlagenverzeichnis_rows']:4}  Klasse {info['gitter_anlagenklasse']:3}  "
              f"Bilanzpos {info['gitter_bilanzposition']:3}  Offene Punkte {info['open_points']:3}")
    (a.out / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(summary)} Arbeitsversionen -> {a.out}")


if __name__ == "__main__":
    main()
