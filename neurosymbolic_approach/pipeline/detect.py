"""Step 3b - detection: name every irregularity of the document in the answer keys' vocabulary.

Nothing new is read here; each finding comes from evidence an earlier stage already produced (a glued token, a page
without a header, a subtotal that does not tie, a hidden column ...). Each finding: type, evidence, where.
The `planted` list of a key uses the same type names, so detection can be scored.
"""

from __future__ import annotations

import re
from collections import defaultdict

from .schema import Extracted, Layout, Tokens

SELECTION_RX = re.compile(r"selektion|sortier|datenstatistik|anzeigevariante", re.I)


def detect(tk: Tokens, lay: Layout, ex: Extracted, report: dict, decisions: dict) -> dict:
    found: list[dict] = []

    def add(kind, evidence, where=None):
        found.append({"type": kind, "evidence": evidence, "where": where or []})

    rows = {r.id: r for r in ex.rows}
    bands = {r.id: r for r in lay.rows}
    by_page = defaultdict(list)
    for t in tk.tokens:
        if t.page is not None:
            by_page[t.page].append(t)

    # ---- tokens: amounts printed without a gap
    glued = [t for t in tk.tokens if t.glued_from and t.id.endswith("a")]
    suspect = sorted({rid for rid, d in decisions.items() for v in d.get("flags", {}).values() if v == "glued_suspect"})
    if glued or suspect:
        add("glued_numbers", f"{len(glued)} printed string(s) hold two amounts; {len(suspect)} row(s) where the laws "
                             f"suggest two amounts were read as one",
            [{"page": t.page, "printed": t.glued_from} for t in glued] + [{"row": rid, "pointer": rows[rid].pointer} for rid in suspect])

    # ---- pages that hold no schedule rows
    if tk.format in ("pdf", "ocr-pdf"):
        sched_pages = {p for s in lay.sections for p in s.pages}
        first_sched = min(sched_pages) if sched_pages else None
        for pg, st in sorted(lay.page_status.items(), key=lambda kv: int(kv[0])):
            words = by_page.get(int(pg), [])
            text = " ".join(t.text for t in words)
            if st == "empty":
                add("noise_page_blank", "page has no printed text", [{"page": int(pg)}])
            elif st == "no_amounts" and SELECTION_RX.search(text):
                continue  # report selection page of an ERP list: part of the print, not noise
            elif st == "no_amounts":
                kind = "noise_page_cover" if first_sched is None or int(pg) < first_sched else "noise_page_unrelated"
                add(kind, f"page has text but no amounts ({len(words)} words)", [{"page": int(pg)}])
            elif st == "no_header":
                add("noise_page_unrelated", "page has amounts but no schedule header and does not continue one",
                    [{"page": int(pg)}])
        cont = [int(p) for p, st in lay.page_status.items() if st == "continues"]
        if cont:
            add("header_only_on_first_page", f"{len(cont)} page(s) carry rows but no column header", [{"page": p} for p in cont])

    # ---- periods
    years = sorted({s.period_end[:4] for s in lay.sections if s.period_end})
    if len(years) >= 2:
        add("multi_year", f"sections for {', '.join(years)}",
            [{"section": s.id, "period_end": s.period_end} for s in lay.sections])
    if any(s.unit == "TEUR" for s in lay.sections):
        add("teur_rounding", "amounts are printed in thousand euro; sums may differ by rounding")

    # ---- rows
    wrapped = [r for r in lay.rows if r.wrapped]
    if wrapped:
        add("wrapped_description", f"{len(wrapped)} description(s) continue on the next line",
            [{"row": r.id, "page": r.page, "label": r.label} for r in wrapped])
    carry = [rid for rid, d in decisions.items() if d.get("kind") == "carry_over"]
    if carry:
        add("carry_over", f"{len(carry)} carry-over line(s) repeat a running sum",
            [{"row": rid, "pointer": rows[rid].pointer} for rid in carry])
    nt = [(rid, f) for rid, d in decisions.items() for f, v in d.get("flags", {}).items()
          if v == "non_tying" or (v == "sum_unverified" and d.get("kind") == "subtotal")]
    if nt:
        add("non_tying_subtotal", f"{len(nt)} printed total value(s) differ from the sum of their rows",
            [{"row": rid, "field": f, "pointer": rows[rid].pointer} for rid, f in nt])
    tf = [(rid, f) for rid, d in decisions.items() for f, v in d.get("flags", {}).items() if v in ("transfer_incomplete", "law_implies_value")]
    if tf:
        add("transfer_breaks_cost_law", "a transfer is missing from a row although its cost at year end includes it",
            [{"row": rid, "field": f, "pointer": rows[rid].pointer} for rid, f in tf])
    after = [rid for rid, d in decisions.items() if "below the grand total" in d.get("why", "")]
    multi = [rid for rid in after if sum(1 for c in rows[rid].cells.values() if c.status == "value") >= 3]
    if len(multi) >= 2:
        add("unrelated_table", f"{len(multi)} rows with several amounts below the grand total",
            [{"row": rid, "label": rows[rid].label} for rid in multi])

    # ---- cells
    # clipped column: numbers without their decimals, all ending at the same x, right of everything else
    cut = []
    for pg, words in by_page.items():
        if not any(t.kind == "amount" for t in words):
            continue
        edge = max(t.frame[2] for t in words if t.kind != "int" or len(t.text) > 1)
        clipped = [t for t in words if abs(t.frame[2] - edge) <= 0.6 and re.fullmatch(r"\d{1,3}(\.\d{3})*", t.text)]
        if sum(1 for t in clipped if t.text != "0") >= 3:
            cut += [{"page": pg, "printed": t.text} for t in clipped]
    if cut:
        add("column_cut_off", f"{len(cut)} number(s) run past the page edge and lose characters", cut[:20])
    if tk.format == "xlsx":
        amount_cols = {(s.sheet, c.col) for s in lay.sections for c in s.columns}
        as_text = [t for t in tk.tokens if t.kind == "amount" and (t.sheet, t.col) in amount_cols]
        if as_text:
            add("numbers_as_text", f"{len(as_text)} amount cell(s) hold text, not numbers",
                [{"cell": t.id, "printed": t.text} for t in as_text[:20]])
        hidden = defaultdict(int)
        for t in tk.tokens:
            if t.hidden_col and t.kind in ("number", "amount") and t.sheet is not None:
                hidden[(t.sheet, t.col)] += 1
        hidden_cols = [(sh, c) for (sh, c), n in hidden.items() if n >= 2]
        if hidden_cols:
            add("hidden_column", "a hidden column holds numbers; it is not part of the schedule",
                [{"sheet": sh, "col": c} for sh, c in hidden_cols])
    for s in report.get("sections", []):
        if s["gate"] != "ok":
            add("mapping_unverified", s["gate"], [{"section": s["id"]}])
    if lay.abstain:
        add("not_read", lay.abstain)
    variant = "problem" if any(f["type"] not in ("not_read", "mapping_unverified") for f in found) else "perfect"
    return {"doc_id": lay.doc_id, "variant_guess": variant if not lay.abstain else "unknown", "findings": found}


def score_detection(det: dict, key: dict, readable: bool) -> dict:
    """Type-level: which planted problem types did we name, which did we miss, which did we invent."""
    planted = {p["type"] for p in key.get("planted", [])}
    named = {f["type"] for f in det["findings"]} - {"not_read", "mapping_unverified"}
    return {"planted": sorted(planted), "detected": sorted(named), "hit": sorted(planted & named),
            "missed": sorted(planted - named), "false_alarm": sorted(named - planted), "readable": readable,
            "variant_key": key["meta"].get("variant"), "variant_guess": det["variant_guess"]}
