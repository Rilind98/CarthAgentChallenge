"""Step 3 - symbolic verification. Nothing here reads the page; it reasons over the extracted cells.

1. Mapping: every column has candidate fields (step 1). The hypothesis whose fields make the accounting laws hold on
   the rows wins. A swapped or shifted column breaks the laws, so it loses.
2. Row laws: each row is checked with every law its printed columns allow (derived by elimination).
3. Totals: a total is a row equal to the sum of the not-yet-summed rows above it, labelled or not. A labelled total
   that differs from its rows by one field is the client's broken sum: kept as printed, flagged (rule 05).
4. Table end: rows after the grand total (reconciliations, forecasts) are not part of the schedule.
5. Decision per value: emit, emit + flag, or blank + flag.
"""

from __future__ import annotations

import itertools
from collections import defaultdict

import os

from .lexicon import checkable_laws, law_name
from .numbers import field_value
from .schema import Extracted, Layout, Tokens


def _tol(law: dict, teur: bool) -> int:
    return 50_000 * len(law) if teur else 0


def _vals(row, mapping) -> dict:
    out = {}
    for cid, f in mapping.items():
        if f:
            c = row.cells[cid]
            out[f] = None if c.cents is None else field_value(f, c.cents)
    return out


def _law_ok(vals: dict, law: dict, tol: int):
    if any(vals.get(f) is None for f in law):
        return None
    return abs(sum(c * vals[f] for f, c in law.items())) <= tol


def select_mapping(sec, rows):
    options = []
    for c in sec.columns:
        cands = [(f, s) for f, s, _ in c.candidates][:3]
        options.append(cands or [(None, 0.0)])
    total = 1
    for o in options:
        total *= len(o)
    if total > 20000:  # keep only the best candidate where the header is unambiguous
        options = [o if len(o) == 1 or o[0][1] - o[1][1] < 0.3 else o[:1] for o in options]
    sample = rows[:300]
    teur = sec.unit == "TEUR"
    results = []
    for combo in itertools.product(*options):
        fields = [f for f, _ in combo if f]
        if len(fields) != len(set(fields)):
            continue
        mapping = {c.id: f for c, (f, _) in zip(sec.columns, combo)}
        laws = checkable_laws(set(fields))
        passed = failed = 0
        for r in sample:
            v = _vals(r, mapping)
            for law in laws:
                ok = _law_ok(v, law, _tol(law, teur))
                if ok is True:
                    passed += 1
                elif ok is False:
                    failed += 1
        lex = sum(s for _, s in combo)
        score = lex + 2 * (passed - 2 * failed) / max(1, len(sample))
        results.append({"score": round(score, 3), "lexicon": round(lex, 3), "laws": len(laws), "passed": passed,
                        "failed": failed, "mapping": mapping})
    if not results:  # every hypothesis reuses a field: let the weakest column go unmapped
        live = [i for i in range(len(options)) if options[i][0][0] is not None]
        if not live:
            return {c.id: None for c in sec.columns}, []
        weakest = min(live, key=lambda i: options[i][0][1])
        options[weakest] = [(None, 0.0)]
        sec2 = sec.model_copy(update={"columns": [c if i != weakest else c.model_copy(update={"candidates": []})
                                                  for i, c in enumerate(sec.columns)]})
        return select_mapping(sec2, rows)
    results.sort(key=lambda x: -x["score"])
    return results[0]["mapping"], results[:4]


def _match(items, vals, tol, min_items, allow_one_off):
    """Walk up over the not-yet-summed rows; first k whose sum equals vals. Returns (k, fields_off) or None."""
    open_items = [it for it in items if not it["consumed"]]
    fields = [f for f, v in vals.items() if v is not None]
    if not fields:
        return None
    acc = defaultdict(int)
    unknown = set()  # a field some summed row could not show: not comparable
    near = None
    for k, it in enumerate(reversed(open_items), 1):
        for f in fields:
            if f in it["vals"]:
                acc[f] += it["vals"][f]
            else:
                unknown.add(f)
        if k < min_items:
            continue
        tol_k = tol * (k + 1)  # thousand-euro figures are rounded one by one
        comp = [f for f in fields if f not in unknown]
        if not comp:
            continue
        off = [f for f in comp if abs(acc[f] - vals[f]) > tol_k]
        same_nonzero = sum(1 for f in comp if f not in off and vals[f] != 0)
        if not off and (same_nonzero >= 1 or all(v == 0 for v in acc.values())):
            return k, []
        if allow_one_off and near is None and len(off) == 1 and len(comp) >= 4 and same_nonzero >= 2:
            near = (k, off)
    return near


def _totals(rows, mapping, tol1):
    """Subtotals by arithmetic: a total equals the sum of the not-yet-summed rows before it (in reading order)."""
    items, kinds = [], {}
    for r in rows:
        v = {f: x for f, x in _vals(r, mapping).items() if x is not None}
        if r.kind_guess == "carry_over":
            kinds[r.id] = {"kind": "carry_over", "level": 0, "sums": [], "flags": {}, "why": "carry-over label"}
            continue
        m = None
        if r.kind_guess == "asset":
            if sum(1 for x in v.values() if x) >= 3:
                m = _match(items, v, tol1, 2, False)
            if not m:
                items.append({"id": r.id, "level": 0, "vals": v, "consumed": False})
                kinds[r.id] = {"kind": "asset", "level": 0, "sums": [], "flags": {}, "why": "leaf row"}
                continue
            why = "no total label, but equals the sum of the rows above"
        else:
            m = _match(items, v, tol1, 1, True)
            why = "label says total; equals the sum of the rows above" if m else ""
        if m:
            k, off = m
            open_items = [it for it in items if not it["consumed"]][-k:]
            for it in open_items:
                it["consumed"] = True
            level = 1 + max(it["level"] for it in open_items)
            flags = {f: "non_tying" for f in off}
            closes_all = not any(not it["consumed"] for it in items) and (k >= 2 or level >= 2)
            items.append({"id": r.id, "level": level, "vals": v, "consumed": False})
            kinds[r.id] = {"kind": "subtotal", "level": level, "sums": [it["id"] for it in open_items], "flags": flags,
                           "closes_all": closes_all, "why": why + (f"; client sum off in {off}" if off else "")}
        elif r.label.strip():
            kinds[r.id] = {"kind": "subtotal", "level": 1, "sums": [], "flags": {"*": "sum_unverified"},
                           "why": "label says total, but no run of rows above adds up to it"}
        else:
            kinds[r.id] = {"kind": "dropped", "level": 0, "sums": [], "flags": {}, "why": "unlabelled and sums nothing"}
    return kinds


def verify(tk: Tokens, lay: Layout, ex: Extracted) -> tuple[dict, dict]:
    report = {"doc_id": lay.doc_id, "abstain": lay.abstain, "sections": [], "completeness": {}}
    decisions: dict[str, dict] = {}
    used_tokens = {tid for r in ex.rows for c in r.cells.values() for tid in c.tokens}
    ocr = tk.format == "ocr-pdf"
    min_conf = float(os.environ.get("OCR_MIN_CONF", "0.75"))  # tune later on the confidence in 05_values.json
    for sec in lay.sections:
        rows = [r for r in ex.rows if r.section == sec.id]
        if not rows:
            continue
        teur = sec.unit == "TEUR"
        tol1 = 50_000 if teur else 0
        mapping, alternatives = select_mapping(sec, rows)
        observed = {f for f in mapping.values() if f}
        laws = checkable_laws(observed)
        covered = {f for law in laws for f in law}
        # ---- 2. row laws
        row_laws, checked, passed = {}, 0, 0
        for r in rows:
            v = _vals(r, mapping)
            res = []
            for law in laws:
                ok = _law_ok(v, law, _tol(law, teur))
                if ok is not None:
                    checked += 1
                    passed += ok
                    res.append((law_name(law), ok))
            row_laws[r.id] = res
        rate = passed / checked if checked else None
        gate = "ok" if rate is None or rate >= 0.6 else f"abstain: laws hold on only {rate:.0%} of checks"
        # support for each chosen field: a law that held with a non-zero value, or an unambiguous header
        law_hits = defaultdict(int)
        for r in rows:
            v = _vals(r, mapping)
            for law in laws:
                if _law_ok(v, law, _tol(law, teur)):
                    for f in law:
                        law_hits[f] += bool(v.get(f))
        support = {}
        for c in sec.columns:
            f = mapping.get(c.id)
            if not f:
                continue
            sc = {g: x for g, x, _ in c.candidates}
            rival = max([x for g, x in sc.items() if g != f], default=0.0)
            why = next((w for g, x, w in c.candidates if g == f), "")
            if law_hits[f] and rate is not None and rate >= 0.6:
                support[c.id] = "law"
            elif why.startswith("vlm; lexicon says"):
                support[c.id] = "weak"  # no law can decide and the two readers disagree: leave it blank
            elif sc.get(f, 0) >= 0.5 and sc.get(f, 0) - rival >= 0.099:
                support[c.id] = "header"
            else:
                support[c.id] = "weak"
        # ---- 3. totals by arithmetic, top-down and bottom-up (some prints put the total above its rows)
        kinds = _totals(rows, mapping, tol1)
        reverse = False
        unverified = sum(1 for k in kinds.values() if k["flags"].get("*") == "sum_unverified")
        if unverified:
            back = _totals(list(reversed(rows)), mapping, tol1)
            confirmed = lambda ks: sum(1 for k in ks.values() if k["kind"] == "subtotal" and k["sums"])
            if confirmed(back) > confirmed(kinds):
                kinds, reverse = back, True
                for k in kinds.values():
                    k["why"] = k["why"].replace("above", "below")
                    k["closes_all"] = False  # no table-end logic when totals sit above their rows
        # ---- 4. grand total and table end (per page / sheet)
        order = [r for r in rows if kinds[r.id]["kind"] != "dropped"]
        grand_ids = []
        for i, r in enumerate(order):
            kd = kinds[r.id]
            if kd["kind"] == "subtotal" and kd.get("closes_all"):
                gv = _vals(r, mapping)
                later = any(kinds[o.id]["kind"] == "subtotal" and kinds[o.id]["sums"] and _vals(o, mapping) != gv
                            for o in order[i + 1:])
                if not later:
                    kd["kind"], kd["why"] = "grand_total", "closes every row above it and no total follows"
                    grand_ids.append(r.id)
        if grand_ids:
            last = next(r for r in order if r.id == grand_ids[-1])
            gv = _vals(last, mapping)
            after = order[order.index(last) + 1:]
            for o in after:
                if (o.pointer.get("page"), o.pointer.get("sheet")) != (last.pointer.get("page"), last.pointer.get("sheet")):
                    continue
                if _vals(o, mapping) == gv:
                    kinds[o.id].update(kind="grand_total", why="repeats the grand total")
                else:
                    kinds[o.id].update(kind="dropped", why="below the grand total: not part of the schedule")
            # a total printed before the table (auditor's SUBTOTAL row) is grand if it equals all assets
            all_assets = defaultdict(int)
            for o in order:
                if kinds[o.id]["kind"] == "asset":
                    for f, x in _vals(o, mapping).items():
                        all_assets[f] += x or 0
            for o in order:
                if kinds[o.id]["kind"] == "subtotal" and kinds[o.id]["flags"].get("*") == "sum_unverified":
                    v = {f: x for f, x in _vals(o, mapping).items() if x is not None}
                    if v and all(abs(all_assets.get(f, 0) - x) <= tol1 * 20 for f, x in v.items()):
                        kinds[o.id].update(kind="grand_total", flags={}, why="equals the sum of all asset rows")
        # ---- transfers net to zero over the register. When they do not, and a row's cost roll-forward fails by exactly
        # that imbalance, both laws point at the same cell: that row's transfer is not what the export shows.
        transfer_fix, imbalance = {}, None
        ucol = next((cid for cid, f in mapping.items() if f == "umbuchung"), None)
        if ucol:
            imbalance = sum(field_value("umbuchung", r.cells[ucol].cents) for r in rows
                            if kinds[r.id]["kind"] == "asset" and r.cells[ucol].cents is not None)
            if imbalance:
                for r in rows:
                    v = _vals(r, mapping)
                    for law in laws:
                        if "umbuchung" in law and _law_ok(v, law, _tol(law, teur)) is False:
                            res = sum(c * v[f] for f, c in law.items())
                            delta = -res / law["umbuchung"]
                            if abs(delta + imbalance) <= tol1:
                                transfer_fix[r.id] = int(v["umbuchung"] + delta)
        # ---- 5. decisions
        for r in rows:
            kd = kinds[r.id]
            if kd["kind"] == "dropped":
                decisions[r.id] = {"emit": False, **kd}
                continue
            fields, printed, flags = {}, {}, dict(kd["flags"])
            fails = [name for name, ok in row_laws[r.id] if not ok]
            if fails:
                flags["*"] = flags.get("*") or "law_fail: " + "; ".join(fails)
            # the failing laws point at an empty cell: either a value went missing (blank it rather than claim zero),
            # or two values were printed without a gap and read as one in the next column (blank that one too)
            suspects = {}
            bad_law_fields = [set(law) for law in laws if _law_ok(_vals(r, mapping), law, _tol(law, teur)) is False] if fails else []
            if fails:
                v = _vals(r, mapping)
                bad_fields = {f for law in laws if _law_ok(v, law, _tol(law, teur)) is False for f in law}
                order = [c.id for c in sec.columns]
                for cid, f in mapping.items():
                    if f in bad_fields and r.cells[cid].status == "empty":
                        suspects[f] = "law_implies_value"
                        i = order.index(cid)
                        nxt = order[i + 1] if i + 1 < len(order) else None
                        if nxt and mapping.get(nxt) in bad_fields and r.cells[nxt].status == "value":
                            suspects[f] = suspects[mapping[nxt]] = "glued_suspect"
            for cid, f in mapping.items():
                if not f:
                    continue
                c = r.cells[cid]
                printed[f] = c.text
                if gate != "ok":
                    fields[f], flags[f] = None, "mapping_unverified"
                elif support.get(cid) == "weak":
                    fields[f], flags[f] = None, "field_uncertain"
                elif f in suspects:
                    fields[f], flags[f] = None, suspects[f]
                elif ocr and c.cents is not None and c.tokens and c.conf < min_conf:
                    fields[f], flags[f] = None, f"low_ocr_conf:{c.conf}"
                elif ocr and fails and any(f in law_fields for law_fields in bad_law_fields):
                    fields[f], flags[f] = None, "ocr_law_fail"  # OCR may have misread: do not trust a row the laws reject
                elif f == "umbuchung" and r.id in transfer_fix:
                    fields[f], flags[f] = None, "transfer_incomplete"
                elif c.cents is None:
                    fields[f], flags[f] = None, c.status
                else:
                    fields[f] = field_value(f, c.cents)
            decisions[r.id] = {"emit": True, **kd, "fields": fields, "printed": printed, "flags": flags,
                               "laws": row_laws[r.id]}
        stats = defaultdict(int)
        for r in rows:
            stats[kinds[r.id]["kind"]] += 1
        report["sections"].append({
            "id": sec.id, "pages": sec.pages, "sheet": sec.sheet, "period_end": sec.period_end, "unit": sec.unit,
            "mapping": {cid: f for cid, f in mapping.items()},
            "mapping_headers": {c.id: c.header for c in sec.columns},
            "field_support": support,
            "transfer_imbalance": imbalance,
            "totals_direction": "bottom-up" if reverse else "top-down",
            "transfer_fixes": {rid: {"printed": r.cells[ucol].text, "implied_cents": v} for rid, v in transfer_fix.items()
                               for r in rows if r.id == rid},
            "hypotheses": alternatives,
            "laws_checked": [law_name(l) for l in laws],
            "fields_without_law": sorted(observed - covered),
            "row_law_checks": {"checked": checked, "passed": passed, "pass_rate": rate},
            "gate": gate,
            "row_kinds": dict(stats),
            "non_tying": [{"row": rid, "fields": [f for f in k["flags"] if k["flags"][f] == "non_tying"]}
                          for rid, k in kinds.items() if "non_tying" in k["flags"].values()],
            "rows": {rid: {k: v for k, v in kd.items() if k != "flags"} | {"flags": kd["flags"], "laws": row_laws[rid]}
                     for rid, kd in kinds.items()},
        })
    # ---- completeness: amounts inside the table columns that no row picked up
    unassigned = []
    if tk.format in ("pdf", "ocr-pdf"):
        from .layout import in_column
        for sec in lay.sections:
            for t in tk.tokens:
                if t.page in sec.pages and t.kind == "amount" and t.id not in used_tokens \
                        and any(in_column(t, c) for c in sec.columns if any(f for f, _, _ in c.candidates)):
                    unassigned.append({"token": t.id, "page": t.page, "text": t.text})
    else:
        rowset = {(r.pointer["sheet"], r.pointer["row"]) for r in ex.rows}
        for sec in lay.sections:
            cols = {c.col for c in sec.columns}
            for t in tk.tokens:
                if t.sheet == sec.sheet and t.col in cols and t.kind == "number" and (t.sheet, t.row) not in rowset \
                        and not t.hidden and t.text not in sec.header_lines:
                    unassigned.append({"token": t.id, "text": t.text})
    report["completeness"] = {"unassigned_amounts": len(unassigned), "examples": unassigned[:15]}
    return report, decisions
