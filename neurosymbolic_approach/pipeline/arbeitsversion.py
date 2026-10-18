"""Step 4b - fill the auditor's Arbeitsversion (template xlsx) from the key-shaped document.

Anlagenverzeichnis: one row per asset of an asset-level schedule.
Gitter nach Anlagenklasse / nach Bilanzposition: one row per class / line item, the printed subtotals, and the
client's grand total in the row 'Summe laut Mandant'.
Every value written is the extracted figure; its source (page + box, or sheet + cell) is in the cell comment.
Columns the client does not print keep (Verzeichnis) or get (Gitter) the template's formula. Every flag, blank and
detected irregularity becomes a line on 'Offene Punkte'.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment

AV_COLS = {"ahk_open": "W", "zugang": "X", "abgang": "Y", "umbuchung": "Z", "ahk_close": "AA", "kum_afa_open": "AB",
           "afa_gj": "AC", "afa_abgang": "AE", "zuschreibung": "AF", "kum_afa_close": "AH", "bw_close": "AI",
           "bw_prev": "AJ"}
AV_FORMULA_COLS = {"AA", "AH", "AI", "AJ"}  # 'Mandant, sonst Formel'
GT_COLS = {"ahk_open": "D", "zugang": "E", "abgang": "F", "umbuchung": "G", "bew_ahk": "H", "ahk_close": "I",
           "kum_afa_open": "J", "afa_gj": "L", "afa_abgang": "N", "zuschreibung": "O", "bew_wb": "Q",
           "kum_afa_close": "R", "bw_close": "S", "bw_prev": "T"}
GT_FIRST, GT_ROWS, GT_SUM_ROW = 8, 20, 28
METHOD = {"lin": "linear", "linear": "linear", "degr": "degressiv", "degressiv": "degressiv", "geb": "Gebäude-AfA",
          "gwg": "GWG", "sb": "Sofortabschreibung", "k": "keine AfA"}
POS266 = [("A.I.1", "selbst geschaffene gewerbliche schutzrechte"), ("A.I.2", "entgeltlich erworbene konzessionen"),
          ("A.I.3", "geschäfts- oder firmenwert"), ("A.I", "immaterielle vermögensgegenstände"),
          ("A.II.1", "grundstücke"), ("A.II.2", "technische anlagen und maschinen"),
          ("A.II.3", "betriebs- und geschäftsausstattung"), ("A.II.4", "anlagen im bau"), ("A.II", "sachanlagen"),
          ("A.III.1", "anteile an verbundenen unternehmen"), ("A.III.2", "ausleihungen an verbundene unternehmen"),
          ("A.III.3", "beteiligungen"), ("A.III.4", "ausleihungen an unternehmen, mit denen ein beteiligungsverh"),
          ("A.III.5", "wertpapiere des anlagevermögens"), ("A.III.6", "sonstige ausleihungen"), ("A.III", "finanzanlagen")]


def gliederung(text: str) -> str | None:
    """§ 266 Abs. 2 code from a printed name; None when nothing or more than one position fits."""
    low = re.sub(r"\s+", " ", text.lower())
    hits = [(len(name), code) for code, name in POS266 if name in low]
    if not hits:
        return None
    best = max(h[0] for h in hits)
    codes = {c for n, c in hits if n == best}
    return codes.pop() if len(codes) == 1 else None


def _label_parts(label: str) -> tuple[str | None, str]:
    """'Summe 0027 EDV-Software' -> ('0027', 'EDV-Software'); '*2101200' -> ('2101200', '')."""
    t = re.sub(r"^\W*(summe|summen|zwischensumme|gesamt|total)\b\s*", "", label.strip(), flags=re.I).lstrip("* ")
    m = re.match(r"^(\d[\d.]*)\s*(.*)$", t)
    if m and not re.fullmatch(r"\d\.", m[1]):
        return m[1].rstrip("."), m[2].strip()
    m = re.match(r"^(?:[A-H]\.\s*)?(?:[IVX]{1,4}\.|\d{1,2}\.)\s*(.*)$", t)  # '2. Technische Anlagen ...', 'II. Sachanlagen'
    return None, (m[1] if m else t).strip()


def _source(src: str, ptr: dict, printed: str | None) -> str:
    if ptr.get("sheet"):
        where = f"{ptr['sheet']}!{ptr.get('cell')}"
    else:
        where = f"Seite {ptr.get('page')}, Box {ptr.get('bbox')}"
    return f"Quelle: {src}, {where}" + (f" · gedruckt '{printed}'" if printed else " · leer gedruckt (= 0)")


def _note(ws, ref, text):
    c = Comment(text, "Datenübernahme")
    c.width, c.height = 320, 60
    ws[ref].comment = c


def write(template: Path, out: Path, doc: dict, detections: list[dict]) -> dict:
    wb = openpyxl.load_workbook(template)
    src = Path(doc["meta"].get("source", "")).name
    open_points: list[tuple[str, str, str]] = []
    seen_net = set()
    rows = doc["rows"]
    if not rows:
        open_points.append(("—", src, f"Dokument nicht übernommen: {doc['meta'].get('abstain') or 'keine Zeilen gelesen'}"))
    years = sorted({r["year"] for r in rows if r.get("year")})
    year = years[-1] if years else None
    if len(years) > 1:
        open_points.append(("—", src, f"Die Übergabe enthält die Jahre {', '.join(map(str, years))}; übernommen ist {year}, "
                                     f"die übrigen stehen nur in der JSON-Ausgabe."))
    rows = [r for r in rows if r.get("year") in (year, None)]
    period_end = max((s["period_end"] for s in doc["meta"]["sections"] if s.get("period_end")), default=None)

    # ---- where each row goes, decided per section (the ground-truth convention of the Arbeitsversion):
    # asset-level section: assets -> Verzeichnis; lower subtotal levels -> Gitter/Anlagenklasse 'Anlagenklasse';
    #   the top subtotal level (balance-sheet position) -> Anlagenklasse 'Zwischensumme Bilanzposition' AND
    #   Bilanzposition 'Bilanzposition'.
    # class-level section (accounts / classes as leaves): leaves -> Anlagenklasse 'Anlagenklasse'; subtotals as above.
    # line-item section (leaves are § 266 positions): everything -> Bilanzposition ('Bilanzposition' / 'Zwischensumme').
    # grand total -> 'Gesamtsumme' on every Gitter sheet that has rows.
    sections: dict[str, list] = {}
    for r in rows:
        sections.setdefault(r.get("section") or "", []).append(r)
    av_rows, ga, gb, grand = [], [], [], []
    leaf_level = {s["id"]: s.get("leaf_level") for s in doc["meta"]["sections"]}
    for sid, sec_rows in sections.items():
        leaves = [r for r in sec_rows if r["kind"] == "asset"]
        subs = [r for r in sec_rows if r["kind"] == "subtotal"]
        lvl = leaf_level.get(sid)
        if lvl:  # the VLM read what a leaf row is
            asset_level, line_items = lvl == "asset", lvl == "balance_sheet_item"
        else:  # fallback without a VLM: dated rows are assets, § 266 names are line items
            asset_level = bool(leaves) and sum(1 for r in leaves if r.get("acq_date")) >= 0.5 * len(leaves)
            line_items = not asset_level and bool(leaves) and sum(1 for r in leaves if gliederung(r["label"])) >= 0.5 * len(leaves)
        summed_by_subtotal = {i for r in subs for i in r.get("sums", [])}
        sub_ids = {r["id"] for r in subs}
        parent = {i: r["id"] for r in subs for i in r.get("sums", [])}
        # a position: no other subtotal adds it up, and (in an asset list) it adds up subtotals itself
        positions = {r["id"] for r in subs if r["id"] not in summed_by_subtotal and (
            not asset_level or any(i in sub_ids for i in r.get("sums", [])))}
        candidates = [r for r in sec_rows if r["kind"] == "subtotal" and r["id"] not in positions] + \
            ([] if asset_level else leaves)
        if positions:
            class_rows = {r["id"] for r in candidates if parent.get(r["id"]) in positions}
        else:  # no position level: the highest rows below the grand total
            class_rows = {r["id"] for r in candidates if r["id"] not in parent}
        for r in sec_rows:
            if r["kind"] == "grand_total":
                grand.append(r)
            elif r["kind"] == "carry_over":
                continue
            elif line_items:
                gb.append((r, "Bilanzposition" if r["kind"] == "asset" else "Zwischensumme Bilanzposition"))
            elif r["kind"] == "asset" and asset_level:
                av_rows.append(r)
            elif r["id"] in positions:
                ga.append((r, "Zwischensumme Bilanzposition"))
                gb.append((r, "Bilanzposition"))
            elif r["id"] in class_rows:  # the row directly under a position (or the top level when there is none)
                ga.append((r, "Anlagenklasse"))

    # ---------------------------------------------------------------- Anlagenverzeichnis
    av = wb["Anlagenverzeichnis"]
    n_av = 0
    for r in av_rows:
        n_av += 1
        i = n_av + 1
        no = r.get("asset_no") or _label_parts(r["label"])[0] or ""
        sub = str(r.get("sub_no") or "")
        av[f"A{i}"] = no + (f"/{sub}" if sub and sub != "0" else "")  # '100/a', '1/1': as the Arbeitsversion writes it
        av[f"B{i}"] = r.get("name") or _label_parts(r["label"])[1]
        ctx = [c for c in r.get("context", []) if c]
        if ctx:
            code, name = _label_parts(re.sub(r"^konto\s*", "", ctx[-1], flags=re.I))
            av[f"C{i}"] = name or ctx[-1]
            if code:
                av[f"H{i}"] = code
            g = next((gliederung(c) for c in reversed(ctx) if gliederung(c)), None)
            if g:
                av[f"E{i}"] = g
        if r.get("acq_date"):
            try:
                av[f"K{i}"] = date.fromisoformat(r["acq_date"])
            except ValueError:
                pass
        if r.get("method"):
            av[f"P{i}"] = METHOD.get(r["method"].lower(), r["method"])
        for f, col in AV_COLS.items():
            ref = f"{col}{i}"
            if f in r["fields"]:
                v = r["fields"][f]
                if v is None:
                    av[ref] = None
                    open_points.append(("Anlagenverzeichnis", ref, f"{f} von Anlage {av[f'A{i}'].value} leer gelassen: "
                                        f"{r['flags'].get(f) or r['flags'].get('*')}"))
                else:
                    av[ref] = v / 100
                    _note(av, ref, _source(src, r["pointer"]["fields"].get(f, r["pointer"]), r["printed"].get(f)))
            elif col not in AV_FORMULA_COLS:
                av[ref] = None
        for f, flag in r["flags"].items():
            if flag in ("non_tying", "sum_unverified") or flag.startswith("law_fail"):
                open_points.append(("Anlagenverzeichnis", f"A{i}", f"Anlage {av[f'A{i}'].value}: {flag}"))
    if any("abgang_bw" in r["fields"] for r in av_rows):
        open_points.append(("Anlagenverzeichnis", "Y", "Abgang nur zum Restbuchwert gedruckt (Nettodarstellung): "
                                                     "Bruttodarstellung nach § 284 Abs. 3 HGB nachfordern."))

    # ---------------------------------------------------------------- Gitter (both sheets)
    for name in ("Gitter nach Bilanzposition", "Gitter nach Anlagenklasse"):
        if period_end:
            wb[name]["C2"] = date.fromisoformat(period_end)

    def fill(gt, gname, i, r, zeilenart):
        code, name = _label_parts(r["label"])
        g = gliederung(r["label"]) or next((gliederung(c) for c in reversed(r.get("context", [])) if gliederung(c)), None)
        if g:
            gt[f"A{i}"] = g
        if code:
            gt[f"B{i}"] = code
        gt[f"C{i}"] = name or r["label"] or zeilenart
        gt[f"V{i}"] = zeilenart
        if gname == "Gitter nach Anlagenklasse" and r.get("context"):
            pos = next((c for c in reversed(r["context"]) if gliederung(c)), None)
            if pos:
                gt[f"W{i}"] = pos
        for f, col in GT_COLS.items():
            if f not in r["fields"]:
                continue
            ref, v = f"{col}{i}", r["fields"][f]
            if v is None:
                open_points.append((gname, ref, f"{f} in Zeile '{gt[f'C{i}'].value}' leer gelassen: {r['flags'].get(f) or r['flags'].get('*')}"))
                continue
            gt[ref] = (-v if f == "bew_wb" else v) / 100  # 'Bewegung netto' of the depreciation is signed: write-backs reduce it
            _note(gt, ref, _source(src, r["pointer"]["fields"].get(f, r["pointer"]), r["printed"].get(f)))
        formulas = {"I": ("D", "E", "F", "G", f"=D{i}+E{i}-F{i}+G{i}"), "R": ("J", "L", "N", "O", f"=J{i}+L{i}-N{i}-O{i}"),
                    "S": ("I", "R", f"=I{i}-R{i}"), "T": ("D", "J", f"=D{i}-J{i}")}
        for col, spec in formulas.items():
            *needs, formula = spec
            if gt[f"{col}{i}"].value is None and all(gt[f"{c}{i}"].value is not None for c in needs):
                gt[f"{col}{i}"] = formula  # 'Mandant, sonst Formel': only when every input of the row is printed
        for f, flag in r["flags"].items():
            if flag in ("non_tying", "sum_unverified"):
                open_points.append((gname, f"C{i}", f"Summe '{gt[f'C{i}'].value}' weicht von ihren Zeilen ab ({f}): "
                                                    "wie gedruckt übernommen."))
        if "abgang_bw" in r["fields"] and gname not in seen_net:
            seen_net.add(gname)
            open_points.append((gname, "F", "Abgang nur zum Restbuchwert gedruckt (Nettodarstellung): "
                                            "Bruttodarstellung nach § 284 Abs. 3 HGB nachfordern."))

    filled = {}
    for gname, body in (("Gitter nach Anlagenklasse", ga), ("Gitter nach Bilanzposition", gb)):
        if not body and not (grand and gname == "Gitter nach Anlagenklasse" and not gb):
            continue
        gt = wb[gname]
        if len(body) > GT_ROWS:
            gt.insert_rows(GT_FIRST + GT_ROWS, amount=len(body) - GT_ROWS)
        sum_row = GT_FIRST + max(GT_ROWS, len(body))
        for k, (r, zeilenart) in enumerate(body):
            fill(gt, gname, GT_FIRST + k, r, zeilenart)
        if grand:
            fill(gt, gname, sum_row, grand[0], "Gesamtsumme")
            gt[f"C{sum_row}"] = "Summe laut Mandant"
        filled[gname] = len(body)
    gname, body = "+".join(filled) or "-", [None] * sum(filled.values())

    # ---------------------------------------------------------------- Offene Punkte
    for d in detections:
        if d["type"] in ("not_read",):
            continue
        open_points.append(("—", src, f"Erkannt: {d['type']} – {d['evidence']}"))
    op = wb["Offene Punkte"]
    if open_points:
        op["A4"] = None
        for k, (tab, where, question) in enumerate(open_points):
            op[f"A{4 + k}"], op[f"B{4 + k}"], op[f"C{4 + k}"], op[f"D{4 + k}"] = tab, where, question, "offen (Datenübernahme)"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"file": str(out), "anlagenverzeichnis_rows": n_av, "gitter_sheet": gname, "gitter_rows": len(body),
            "grand_total_row": bool(grand), "open_points": len(open_points), "year": year}
