"""Step 1 - layout proposal. This is the slot for the neural model (VLM); today a deterministic stand-in fills it.

Output = what a VLM would be asked for: column boxes with their header text and candidate fields, row boxes with a
kind guess and their label, the period and the unit. It never outputs a figure (schema.Layout rejects any amount
in its strings). Which column is which field is NOT decided here: step 3 picks the hypothesis the laws support.
"""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import date
from statistics import median

from . import lexicon as lx
from .numbers import TEUR_RE
from .schema import Column, Layout, RowBand, Section, Token, Tokens

SUM_RX = re.compile(r"^\*|summe|gesamt|total|zwischensumme|^sum\b|^summen", re.I)
CARRY_RX = re.compile(r"^\W*(übertrag|uebertrag|vortrag|transport)\b(?!\w)", re.I)
BLOCK_LABEL_RX = re.compile(r"^(ansch-?/?\s*herst-?k\.?|ahk|absch|abschreibung(en)?|bw|buchwert(e)?|anschaffungskosten)$",
                            re.I)
TYPE_WORDS = [("AHK", re.compile(r"^(ak/hk|ahk|ak|ansch-?/?\s*herst-?k\.?|anschaffungskosten)$")),
              ("AFA", re.compile(r"^(absch|abschreibung(en)?|afa|kum\.?\s*afa)$")),
              ("BW", re.compile(r"^(bw|rbw|buchwert(e)?|restbuchwert)$"))]
FOOTER_RX = re.compile(r"^(seite|datum:|blatt|page)\b|seite\s*:?\s*\d|von\s+\d+$", re.I)
TITLE_RX = re.compile(r"berichtsdatum|erstellungsdatum|stichtag|seite|blatt\b|anlagengitter|anlagespiegel|"
                      r"anlageverzeichnis|entwicklung des|jahresabschluss|wirtschaftsjahr", re.I)
TEUR_RX = re.compile(r"\bteur\b|\bt€|tsd\.?\s*(eur|€)|in\s+tausend", re.I)


def _kind_guess(label: str) -> str:
    if CARRY_RX.search(label):
        return "carry_over"
    if not label.strip() or SUM_RX.search(label.strip()):
        return "total"
    if re.match(r"^(?:[A-H]\.\s*)?[IVX]{1,4}\.\s", label.strip()):  # balance-sheet group (A. II.) carrying amounts
        return "total"
    return "asset"


def _amount_like(t: Token, teur: bool) -> bool:
    """A printed amount. In thousand-euro prints amounts are integers: 49 · 7.185 · 2.739-"""
    return t.kind == "amount" or (teur and t.kind in ("int", "text") and TEUR_RE.match(t.text) is not None)


def _column_evidence(t: Token, teur: bool) -> bool:  # a lone "-" is a printed zero: it marks a column too
    if t.glued_from and t.id.endswith("a"):  # left half of a glued pair: its position is not its column's
        return False
    return _amount_like(t, teur) or t.kind == "dash"


# ================================================================ PDF

def _lines(toks: list[Token]) -> list[dict]:
    toks = sorted(toks, key=lambda t: ((t.frame[1] + t.frame[3]) / 2, t.frame[0]))
    lines: list[dict] = []
    for t in toks:
        cy, h = (t.frame[1] + t.frame[3]) / 2, t.frame[3] - t.frame[1]
        if lines and abs(cy - lines[-1]["cy"]) <= max(1.0, 0.45 * h):
            lines[-1]["toks"].append(t)
        else:
            lines.append({"cy": cy, "toks": [t]})
    for i, ln in enumerate(lines):
        ln["toks"].sort(key=lambda t: t.frame[0])
        ln["top"] = min(t.frame[1] for t in ln["toks"])
        ln["bottom"] = max(t.frame[3] for t in ln["toks"])
        ln["text"] = " ".join(t.text for t in ln["toks"])
        ln["no"] = i + 1
    return lines


def num_right(t: Token) -> float:
    """Right edge of the digits: a trailing minus or bracket sticks out of the column."""
    if t.text.endswith(("-", ")")) and len(t.text) > 1:
        return t.frame[2] - (t.frame[2] - t.frame[0]) / len(t.text)
    return t.frame[2]


def _clusters(data_lines: list[dict], teur: bool) -> list[dict]:
    xs = sorted(((num_right(t), t) for ln in data_lines for t in ln["toks"] if _column_evidence(t, teur)), key=lambda p: p[0])
    cl: list[dict] = []
    for x, t in xs:
        if cl and x - cl[-1]["xs"][-1] <= 3.0:
            cl[-1]["xs"].append(x)
            cl[-1]["toks"].append(t)
        else:
            cl.append({"xs": [x], "toks": [t]})
    need = 2 if len(data_lines) > 2 else 1
    out = []
    for c in cl:
        if len(c["xs"]) >= need:
            out.append({"anchor": round(median(c["xs"]), 2), "x0": min(t.frame[0] for t in c["toks"]),
                        "x1": max(t.frame[2] for t in c["toks"]), "n": len(c["xs"])})
    return out


def in_column(t: Token, col) -> bool:
    """Shared by step 1 and step 2: does a printed token sit in this column box?"""
    anchor, (x0, x1) = (col["anchor"], (col["x0"], col["x1"])) if isinstance(col, dict) else (col.anchor, col.span)
    if abs(num_right(t) - anchor) <= 3.5:
        return True
    cx = (t.frame[0] + t.frame[2]) / 2
    return x0 - 2 <= cx <= x1 + 2


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _camel_split(t: Token) -> list[tuple[str, float, float]]:
    """'ZuschreibungenAbschreibungen' -> two parts with proportional x-ranges (glued header words)."""
    parts = re.split(r"(?<=[a-zäöüß.]{2})(?=[A-ZÄÖÜ][a-zäöüß]{2}|AfA|AHK|BW\b)", t.text)
    if len(parts) < 2:
        return [(t.text, t.frame[0], t.frame[2])]
    out, x, w = [], t.frame[0], (t.frame[2] - t.frame[0]) / max(1, len(t.text))
    for p in parts:
        out.append((p, x, x + w * len(p)))
        x += w * len(p)
    return out


STRONG_ROLE_RX = re.compile("|".join(rx.pattern for r, rx in lx.ROLES if r in ("transfer", "writeup", "disposal", "add", "year")))


def _assign_header(lines: list[dict], cols: list[dict]):
    """Header words -> columns.

    1. A phrase over several columns that only names a group ('kumulierte Abschreibungen') is a group header.
    2. The other words are clustered into labels (words that overlap across lines or sit close on one line).
    3. In each column box, the label over the numbers is the column's own. Another label with schedule vocabulary
       is a column printed entirely blank (returned as a phantom); any other label belongs to a text column.
    """
    per_col = [defaultdict(list) for _ in cols]
    group = [defaultdict(list) for _ in cols]
    last_no = lines[-1]["no"] if lines else None
    words = []
    for ln in lines:
        ws = [w for t in ln["toks"] for w in _camel_split(t)]
        h = ln["toks"][0].frame[3] - ln["toks"][0].frame[1]
        phrases, cur = [], [ws[0]]
        for w in ws[1:]:
            (cur.append(w) if w[1] - cur[-1][2] < 0.5 * h else (phrases.append(cur), cur := [w]))
        phrases.append(cur)
        for ph in phrases:
            x0, x1 = ph[0][1], ph[-1][2]
            text = " ".join(w[0] for w in ph)
            hit = [i for i, c in enumerate(cols) if _overlap(x0, x1, c["r0"], c["r1"]) >= min(0.25 * (x1 - x0), 0.5 * (c["r1"] - c["r0"]))]
            low = lx.norm(text)
            n_groups = sum(1 for rx in lx.GROUPS.values() if rx.search(low))
            whole = all(any(t.text == w[0] for t in ln["toks"]) for w in ph)  # no glued word inside
            if len(hit) >= 2 and n_groups <= 1 and not STRONG_ROLE_RX.search(low) and lx.candidates(text) != [] \
                    and ln["no"] != last_no and whole:
                for i in hit:
                    group[i][ln["no"]].append(text)
                continue
            words += [{"text": w[0], "x0": w[1], "x1": w[2], "no": ln["no"], "h": h} for w in ph]
    # words of text-column headers (date, currency, method, useful life ...) go, with the words stacked on them
    textcol = [w for w in words if lx.TEXT_COLUMN.search(lx.norm(w["text"])) and not any(
        rx.search(lx.norm(w["text"])) for _, rx in lx.ROLES[:4])]
    words = [w for w in words if w not in textcol and not any(
        u["no"] != w["no"] and _overlap(w["x0"], w["x1"], u["x0"], u["x1"]) > 0.5 * min(w["x1"] - w["x0"], u["x1"] - u["x0"])
        for u in textcol)]
    # words -> column boxes; inside a box, words that stack across lines or touch on one line form a label
    in_box = defaultdict(list)
    for w in words:
        cx = (w["x0"] + w["x1"]) / 2
        i = next((i for i, c in enumerate(cols) if c["r0"] <= cx <= c["r1"]), None)
        if i is not None:
            in_box[i].append(w)
    phantoms = []
    for i, ws in in_box.items():
        c = cols[i]
        labels: list[list[dict]] = []
        for w in sorted(ws, key=lambda w: w["x0"]):
            cw = (w["x1"] - w["x0"]) / max(1, len(w["text"]))
            hits = [lab for lab in labels if any(
                _overlap(w["x0"], w["x1"], u["x0"], u["x1"]) > 0 or
                (u["no"] == w["no"] and -0.5 < w["x0"] - u["x1"] < 1.3 * max(cw, (u["x1"] - u["x0"]) / max(1, len(u["text"]))))
                for u in lab)]
            if hits:
                merged = [u for lab in hits for u in lab] + [w]
                labels = [lab for lab in labels if not any(lab is h for h in hits)] + [merged]
            else:
                labels.append([w])

        def text_of(lab):
            return " ".join(u["text"] for u in sorted(lab, key=lambda u: (u["no"], u["x0"])))

        def span_of(lab):
            return min(u["x0"] for u in lab), max(u["x1"] for u in lab)

        own = max(labels, key=lambda lab: (_overlap(*span_of(lab), c["x0"], c["anchor"]), span_of(lab)[1]))
        keep = list(own)
        for lab in labels:
            if lab is own:
                continue
            t = text_of(lab)
            cands = lx.candidates(t)
            low = lx.norm(t)
            if not c.get("phantom") and cands and cands[0][1] >= 0.9 and _overlap(*span_of(lab), c["x0"], c["anchor"]) <= 0:
                x0, x1 = span_of(lab)
                phantoms.append({"anchor": round(x1, 2), "x0": x0, "x1": x1, "n": 0, "phantom": True, "label": t})
            elif lx.HARD_EXCLUDE.search(low) or (lx.EXCLUDE.search(low) and not cands):
                continue  # header of a text column (date, method, useful life, cost centre)
            else:
                keep += lab
        for u in sorted(keep, key=lambda u: (u["no"], u["x0"])):
            per_col[i][u["no"]].append(u["text"])
    return per_col, group, phantoms


def _regions(cols: list[dict], lefts: list[float]) -> None:
    """Header region of a column: its whole box, from the right edge of whatever is printed left of it."""
    for i, c in enumerate(cols):
        wall = max([x for x in lefts if x <= c["x0"] + 0.5] + ([cols[i - 1]["anchor"]] if i else []),
                   default=c["x0"] - (c["anchor"] - c["x0"]))
        c["r0"], c["r1"] = wall + 1, c["anchor"] + 2


def _block_by_data(data: list[dict], teur: bool) -> str | None:
    """One record printed over several lines shows in the data: lines that name their value type (AK/HK, kum. AfA,
    RBW), or amount lines without any label under a separate code line."""
    typed, groups, unlabeled = 0, set(), 0
    for ln in data:
        words = [t.text for t in ln["toks"] if not _amount_like(t, teur) and t.kind not in ("dash", "amount")
                 and re.search(r"[A-Za-zÄÖÜäöüß]", t.text)]
        cand = [w.lower() for w in words] + [f"{a} {b}".lower() for a, b in zip(words, words[1:])]
        hit = {g for g, rx in TYPE_WORDS for w in cand if rx.match(w)}
        typed += bool(hit)
        groups |= hit
        unlabeled += not words
    if data and typed >= 0.3 * len(data) and len(groups) >= 2:
        return f"{typed} of {len(data)} amount lines name their value type ({', '.join(sorted(groups))})"
    if len(data) >= 6 and unlabeled >= 0.5 * len(data):
        return f"{unlabeled} of {len(data)} amount lines carry no label (identity printed on another line)"
    return None


def _is_block(per_col_lines: list[dict], groups_text: list[dict]) -> bool:
    """Several header lines, each naming a different group across the same columns: one record = several lines."""
    by_line = defaultdict(lambda: defaultdict(int))
    col_groups = defaultdict(set)
    for ci, lines in enumerate(per_col_lines):
        merged = defaultdict(list, {k: list(v) for k, v in lines.items()})
        for no, words in (groups_text[ci].items() if groups_text else []):
            merged[no] += words
        for no, words in merged.items():
            h = lx.norm(" ".join(words))
            for g, rx in lx.GROUPS.items():
                if rx.search(h):
                    by_line[no][g] += 1
                    col_groups[ci].add((no, g))
    dominant = {no: max(c, key=c.get) for no, c in by_line.items() if max(c.values()) >= 2}
    multi = sum(1 for s in col_groups.values() if len({g for _, g in s}) >= 2 and len({n for n, _ in s}) >= 2)
    return len(set(dominant.values())) >= 2 and multi >= 2


def pdf_layout(tk: Tokens) -> Layout:
    by_page = defaultdict(list)
    for t in tk.tokens:
        by_page[t.page].append(t)
    sections: list[Section] = []
    rows: list[RowBand] = []
    notes: list[str] = []
    cur: Section | None = None
    cur_cols: list[dict] = []
    period: date | None = None
    block_votes = 0
    status: dict[str, str] = {}
    for pg in tk.pages:
        if pg["page"] not in by_page:
            status[str(pg["page"])] = "empty"
    for pno in sorted(by_page):
        lines = _lines(by_page[pno])
        teur = bool(TEUR_RX.search(" ".join(t.text for t in by_page[pno])))  # the unit is read per page
        data = [ln for ln in lines if any(t.kind == "amount" for t in ln["toks"])
                or (teur and sum(_amount_like(t, teur) for t in ln["toks"]) >= 2)]
        if not data:
            notes.append(f"page {pno}: no amounts, skipped")
            status[str(pno)] = "no_amounts"
            continue
        why_block = _block_by_data(data, teur)
        if why_block:
            block_votes += 1
            notes.append(f"page {pno}: {why_block}")
        cols = _clusters(data, teur)
        lefts = [t.frame[2] for ln in data for t in ln["toks"] if not _column_evidence(t, teur)]
        _regions(cols, lefts)
        first = min(data, key=lambda ln: ln["top"])
        # header block: lines just above the first data line that sit over the columns; stop at titles and code lines
        header, last_top = [], None
        for ln in reversed([ln for ln in lines if ln["bottom"] <= first["top"] + 0.5]):
            over = any(_overlap(t.frame[0], t.frame[2], c["r0"], c["r1"]) > 0 for t in ln["toks"] for c in cols)
            codes = sum(1 for t in ln["toks"] if t.kind == "int" and len(t.text) >= 3) >= 2
            if TITLE_RX.search(ln["text"]) or codes:
                if last_top is None:
                    continue
                break
            if over and (last_top is None or last_top - ln["bottom"] < 2.5 * (ln["bottom"] - ln["top"])):
                header.insert(0, ln)
                last_top = ln["top"]
            elif last_top is not None:
                break
        title_lines = [ln["text"] for ln in lines if ln["bottom"] <= first["top"] + 0.5]
        p_end = lx.period_from(title_lines) or period
        if p_end:
            period = p_end
        p_start = lx.period_start_of(p_end) if p_end else None
        hdr_text = [ln["text"] for ln in header]
        if cur is not None and header and hdr_text == cur.header_lines and (p_end and p_end.isoformat()) == cur.period_end:
            cur.pages.append(pno)  # same header and period as the running section: keep its columns
            notes.append(f"page {pno}: repeats the header of section {cur.id}")
            status[str(pno)] = "repeats"
            sec_cols, body_top = cur.columns, header[-1]["bottom"]
            rows_added = _page_rows(lines, by_page, cur, sec_cols, body_top, teur, rows)
            if sum(any(BLOCK_LABEL_RX.match(t.text) for t in ln["toks"]) for ln in data) >= 0.3 * len(data):
                block_votes += 1
            continue
        per_col, groups, phantoms = _assign_header(header, cols) if header else ([{} for _ in cols], [{} for _ in cols], [])
        if phantoms:  # columns printed entirely blank: add them and read the header again
            cols = sorted(cols + phantoms, key=lambda c: c["anchor"])
            _regions(cols, lefts)
            per_col, groups, _ = _assign_header(header, cols)
            notes.append(f"page {pno}: blank column(s) found from the header: {[p['label'] for p in phantoms]}")
        if header and _is_block(per_col, groups):
            block_votes += 1
        columns = []
        for i, c in enumerate(cols):
            htext = " ".join(" ".join(per_col[i][no]) for no in sorted(per_col[i]))
            gtext = " ".join(" ".join(groups[i][no]) for no in sorted(groups[i]))
            cands = lx.candidates(htext, gtext, p_start, p_end)
            amount = re.compile(r"\S*\d[.,]\d\d(?!\d)\S*")  # a misread amount in the header stays out of it
            columns.append(Column(id=f"c{i + 1}", anchor=c["anchor"], span=[round(c["x0"], 2), round(c["x1"], 2)],
                                  header=amount.sub("#", htext), group_header=amount.sub("#", gtext), candidates=cands))
        n_mapped = sum(1 for c in columns if c.candidates)
        if n_mapped >= 2:
            sig = (tuple(round(c.anchor) for c in columns), tuple(tuple(f for f, _, _ in c.candidates) for c in columns), str(p_end))
            status[str(pno)] = "schedule"
            if cur is not None and getattr(cur, "_sig", None) == sig and abs(len(cur.columns) - len(columns)) == 0:
                cur.pages.append(pno)
            else:
                cur = Section(id=f"s{len(sections) + 1}", pages=[pno], period_end=p_end and p_end.isoformat(),
                              period_start=p_start and p_start.isoformat(), unit="TEUR" if teur else "EUR",
                              columns=columns, header_lines=[ln["text"] for ln in header])
                object.__setattr__(cur, "_sig", sig)
                sections.append(cur)
                cur_cols = cols
            body_top = header[-1]["bottom"] if header else 0
        elif cur is not None and cols and sum(any(abs(c["anchor"] - k.anchor) <= 4 for k in cur.columns) for c in cols) >= 0.6 * len(cols):
            cur.pages.append(pno)  # no header on this page: the columns of the previous page carry on
            notes.append(f"page {pno}: no header, continues section {cur.id}")
            status[str(pno)] = "continues"
            body_top = 0
        else:
            notes.append(f"page {pno}: amounts but no schedule header, skipped")
            status[str(pno)] = "no_header"
            continue
        sec_cols = cur.columns
        # data label voting for block layouts that name the line type in the row ("AHK", "Abschreibung", "BW")
        if sum(any(BLOCK_LABEL_RX.match(t.text) for t in ln["toks"]) for ln in data) >= 0.3 * len(data):
            block_votes += 1
        _page_rows(lines, by_page, cur, sec_cols, body_top, teur, rows)
    mode = "block" if block_votes else ("line" if sections else "none")
    abstain = None
    if tk.format == "image-pdf":
        abstain, mode = "no text layer: needs OCR (not wired in yet)", "none"
    elif mode == "block":
        abstain = "multi-line records (one asset/class = several printed lines): block reader not built yet"
    elif not sections:
        abstain = "no schedule header found"
    return Layout(doc_id=tk.doc_id, proposer="heuristic-v0", mode=mode, abstain=abstain, sections=sections,
                  rows=rows if not abstain else [], notes=notes, page_status=status)


def _page_rows(lines, by_page, cur, sec_cols, body_top, teur, rows) -> int:
    """Row bands of one page: data lines become rows, text lines become context or continue a wrapped name."""
    n0 = len(rows)
    pno = lines[0]["toks"][0].page if lines else None
    context: list[str] = []
    last_row: RowBand | None = None
    for ln in lines:
        if ln["top"] < body_top - 0.5:
            continue
        in_cols = [t for t in ln["toks"] if _amount_like(t, teur) and any(in_column(t, c) for c in sec_cols if c.candidates)]
        label_toks = [t for t in ln["toks"] if t.kind not in ("amount", "dash") and t not in in_cols]
        label = " ".join(t.text for t in label_toks)
        if not in_cols:
            if FOOTER_RX.search(ln["text"]) or not label:
                continue
            id_only = last_row is not None and re.fullmatch(r"\d[\w./-]*", last_row.label.strip() or "-") is not None
            if last_row is not None and label_toks and (label_toks[0].frame[0] > _first_x(by_page, last_row) + 4 or id_only) \
                    and ln["top"] - last_row.band[1] < 2.5 * (ln["bottom"] - ln["top"]) and last_row.kind_guess == "asset":
                last_row.label += " " + label  # wrapped description: continuation of the row above
                last_row.wrapped = True
                continue
            context.append(label)
            last_row = None
            continue
        label = re.sub(r"\S*\d[.,]\d\d(?!\d)\S*", "#", label)  # a misread amount stuck to a word stays out of the label
        rb = RowBand(id=f"r{len(rows) + 1}", section=cur.id, page=pno, band=[round(ln["top"], 2), round(ln["bottom"], 2)],
                     kind_guess=_kind_guess(label), label=label, context=[re.sub(r"\S*\d,\d\d\S*", "#", c) for c in context[-3:]])
        rows.append(rb)
        last_row = rb
    return len(rows) - n0


def _first_x(by_page, rb: RowBand) -> float:
    toks = [t for t in by_page[rb.page] if rb.band[0] - 0.5 <= (t.frame[1] + t.frame[3]) / 2 <= rb.band[1] + 0.5]
    return min(t.frame[0] for t in toks) if toks else 0.0


# ================================================================ xlsx

def xlsx_layout(tk: Tokens) -> Layout:
    sections, rows, notes = [], [], []
    block_votes = 0
    for sheet in tk.sheets:
        cells = {(t.row, t.col): t for t in tk.tokens if t.sheet == sheet}
        if not cells:
            continue
        by_row = defaultdict(list)
        for (r, c), t in cells.items():
            by_row[r].append(t)

        def score(r):
            return sum(1 for t in by_row[r] if t.kind == "text" and lx.candidates(t.text))

        scores = {r: score(r) for r in by_row}
        best = max(scores.values(), default=0)
        if best < 2:
            notes.append(f"sheet {sheet}: no schedule header, skipped")
            continue
        heads = []
        for r in sorted(by_row):
            if scores[r] >= max(2, 0.6 * best) and not (heads and r - heads[-1] <= 3):
                heads.append(r)
        if len(heads) > 1:
            notes.append(f"sheet {sheet}: {len(heads)} header rows, one section each")
        bounds = heads[1:] + [max(by_row) + 1]
        prev_end = 0
        for h, nxt in zip(heads, bounds):
            sec_rows = {r: v for r, v in by_row.items() if prev_end < r < nxt}
            block_votes += _xlsx_section(tk, sheet, cells, sec_rows, by_row, h, sections, rows, notes)
            prev_end = max(r for r in sec_rows) if sec_rows else h
    mode = "block" if block_votes else ("line" if sections else "none")
    abstain = None
    if mode == "block":
        abstain = "multi-line records (one asset/class = several rows): block reader not built yet"
    elif not sections:
        abstain = "no schedule header found"
    return Layout(doc_id=tk.doc_id, proposer="heuristic-v0", mode=mode, abstain=abstain, sections=sections,
                  rows=rows if not abstain else [], notes=notes)


def _xlsx_section(tk, sheet, cells, by_row, sheet_rows, h, sections, rows, notes) -> int:
    """One schedule table on a sheet: header block around row h, title rows above it, data rows below."""
    block_votes = 0
    block = [h]
    for step in (-1, 1):
        r = h + step
        while r in by_row and abs(r - h) <= 3 and all(t.kind in ("text", "date") for t in by_row[r]):
            block.append(r)
            r += step
    block.sort()
    all_text = " ".join(t.text for r in range(1, block[-1] + 1) for t in by_row.get(r, []))
    teur = bool(TEUR_RX.search(all_text))
    title, r = [], block[0] - 1  # title rows: text-only rows right above this table's header
    while r >= 1 and len(title) < 8:
        if r in sheet_rows:
            if any(t.kind == "number" for t in sheet_rows[r]):
                break
            title.insert(0, " ".join(t.text for t in sorted(sheet_rows[r], key=lambda t: t.col)))
        r -= 1
    p_end = lx.period_from(title)
    p_start = lx.period_start_of(p_end) if p_end else None
    hidden_cols = {t.col for t in cells.values() if t.hidden_col}
    col_ids = sorted({c for (_, c) in cells} - hidden_cols)
    columns, id_cols, per_line = [], {}, []
    line_groups = [defaultdict(list) for _ in col_ids]
    group_text = defaultdict(list)
    for r in block:
        for t in by_row[r]:
            if t.merged_cols and t.merged_cols[1] > t.merged_cols[0]:
                for c in range(t.merged_cols[0], t.merged_cols[1] + 1):
                    group_text[c].append(t.text)
    for i, c in enumerate(col_ids):
        parts = []
        for r in block:
            t = cells.get((r, c))
            if t and not (t.merged_cols and t.merged_cols[1] > t.merged_cols[0]):
                parts.append(t.text)
                line_groups[i][r].append(t.text)
        htext = " ".join(parts)
        cands = lx.candidates(htext, " ".join(group_text[c]), p_start, p_end)
        idf = lx.id_column(htext) if not cands else None
        if idf and idf not in id_cols:
            id_cols[idf] = c
        if cands:
            columns.append(Column(id=f"{sheet}:{c}", col=c, header=htext, group_header=" ".join(group_text[c]),
                                  candidates=cands))
    if len(columns) < 2:
        notes.append(f"sheet {sheet}: fewer than two amount columns, skipped")
        return 0
    if _is_block(line_groups, []):
        block_votes += 1
    sec = Section(id=f"s{len(sections) + 1}", sheet=sheet, period_end=p_end and p_end.isoformat(),
                  period_start=p_start and p_start.isoformat(), unit="TEUR" if teur else "EUR", columns=columns,
                  id_columns=id_cols, header_lines=[" | ".join(t.text for t in sorted(by_row[r], key=lambda t: t.col)) for r in block])
    sections.append(sec)
    amount_cols = {c.col for c in columns}
    data_rows = []
    context: list[str] = []
    for r in sorted(by_row):
        if r in block:
            continue
        toks = sorted(by_row[r], key=lambda t: t.col)
        vals = [t for t in toks if t.col in amount_cols and t.kind in ("number", "amount", "dash", "uncached")]
        label = " ".join(t.text for t in toks if t.col not in amount_cols and t.kind in ("text", "int")
                         and t.col not in hidden_cols)
        if any(BLOCK_LABEL_RX.match(t.text) for t in toks if t.kind == "text"):
            data_rows.append(r)
        if not any(t.kind in ("number", "amount") for t in vals):
            if label:
                context.append(label)
            continue
        rows.append(RowBand(id=f"r{len(rows) + 1}", section=sec.id, sheet=sheet, row=r,
                            kind_guess=_kind_guess(label), label=label, context=context[-3:]))
    if data_rows and len(data_rows) >= 0.3 * max(1, sum(1 for x in rows if x.section == sec.id)):
        block_votes += 1
    return block_votes


def propose(tk: Tokens) -> Layout:
    return xlsx_layout(tk) if tk.format == "xlsx" else pdf_layout(tk)
