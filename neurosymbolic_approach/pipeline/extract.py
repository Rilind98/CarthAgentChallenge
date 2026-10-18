"""Step 2 - extraction by code. For every row box x column box: copy the printed string, parse it to cents.

This is the only place where figures come into existence. Nothing is decided about which field a column is;
that is step 3. A cell is one of: value, empty (= zero), dash (= zero), ambiguous, unparsed, uncached.
"""

from __future__ import annotations

import re
from collections import defaultdict

from .layout import _amount_like, in_column
from .numbers import fmt_de, parse_amount, parse_date
from .schema import Cell, Extracted, ExtractedRow, Layout, Token, Tokens

METHODS = {"lin", "linear", "geb", "sb", "k", "degr", "degressiv", "gwg", "leist"}


def _union(boxes):
    return [round(min(b[0] for b in boxes), 2), round(min(b[1] for b in boxes), 2),
            round(max(b[2] for b in boxes), 2), round(max(b[3] for b in boxes), 2)]


def _cell(col_id: str, toks: list[Token], teur: bool, style: str = "de") -> Cell:
    c = _cell0(col_id, toks, teur, style)
    if toks:
        c.conf = round(min(t.conf for t in toks), 3)
    return c


def _cell0(col_id: str, toks: list[Token], teur: bool, style: str = "de") -> Cell:
    ids = [t.id for t in toks]
    box = _union([t.bbox for t in toks]) if toks and toks[0].bbox else None
    coord = toks[0].id.split("!")[1] if toks and toks[0].sheet else None
    if not toks:
        return Cell(column=col_id, tokens=[], text="", cents=0, status="empty")
    if len(toks) > 1:
        return Cell(column=col_id, tokens=ids, text=" ".join(t.text for t in toks), cents=None, status="ambiguous", box=box)
    t = toks[0]
    if t.kind == "uncached":
        return Cell(column=col_id, tokens=ids, text=t.text, cents=None, status="uncached", cell=coord)
    if t.kind == "dash":
        return Cell(column=col_id, tokens=ids, text=t.text, cents=0, status="dash", box=box, cell=coord)
    if t.kind == "number":  # xlsx numeric cell: the stored value is the figure
        cents = round(t.number * (100_000 if teur else 100))
        return Cell(column=col_id, tokens=ids, text=t.text, cents=cents, status="value", cell=coord)
    cents = parse_amount(t.text, teur, style)
    if cents is None:
        return Cell(column=col_id, tokens=ids, text=t.text, cents=None, status="unparsed", box=box, cell=coord)
    return Cell(column=col_id, tokens=ids, text=t.text, cents=cents, status="value", box=box, cell=coord)


def _identity_from_label(toks: list[Token], sub_hint: bool) -> dict:
    """asset no / sub no / name / acquisition date / method from the non-amount words of an asset line."""
    ident = {"asset_no": None, "sub_no": None, "name": None, "acq_date": None, "method": None}
    words = [t.text for t in toks]
    if words:
        m = re.match(r"^(\d+)([a-z])?(?=[A-ZÄÖÜ(])(.*)$", words[0])  # '100Warenwirtschaft', '91aProduktionshalle'
        if m:
            ident["asset_no"], ident["sub_no"] = m[1], m[2] or ""
            words = [m[3]] + words[1:]
        elif re.fullmatch(r"\d[\d./-]*[a-z]?", words[0]) and not parse_date(words[0]):
            m2 = re.fullmatch(r"(\d[\d./-]*?)([a-z])?", words[0])
            ident["asset_no"], ident["sub_no"] = m2[1], m2[2] or ""
            words = words[1:]
            if sub_hint and words and re.fullmatch(r"\d{1,3}", words[0]):
                ident["sub_no"] = words.pop(0)
    rest = []
    for w in words:
        d = parse_date(w)
        if d and ident["acq_date"] is None:
            ident["acq_date"] = d.isoformat()
        elif d:
            continue
        elif w.lower() in METHODS and ident["method"] is None:
            ident["method"] = w
        elif w in ("EUR", "TEUR") or (ident["method"] and re.fullmatch(r"\d{1,3}(,\d+)?", w)):
            continue
        else:
            rest.append(w)
    ident["name"] = " ".join(rest) or None
    return ident


def extract(tk: Tokens, lay: Layout) -> Extracted:
    secs = {s.id: s for s in lay.sections}
    by_page = defaultdict(list)
    by_cell = {}
    for t in tk.tokens:
        if t.page is not None:
            by_page[t.page].append(t)
        else:
            by_cell[(t.sheet, t.row, t.col)] = t
    out = []
    for rb in lay.rows:
        sec = secs[rb.section]
        teur = sec.unit == "TEUR"
        year = int(sec.period_end[:4]) if sec.period_end else None
        cells: dict[str, Cell] = {}
        if rb.page is not None:  # PDF: tokens whose centre lies in the row band
            row_toks = [t for t in by_page[rb.page] if rb.band[0] - 0.5 <= (t.frame[1] + t.frame[3]) / 2 <= rb.band[1] + 0.5]
            used = set()
            cand = [t for t in row_toks if _amount_like(t, teur) or t.kind == "dash"]
            lefts = {t.id[:-1]: t for t in cand if t.glued_from and t.id.endswith("a")}
            per_col = {}
            for c in sec.columns:
                toks = [t for t in cand if in_column(t, c) and t.id not in used and not (t.glued_from and t.id.endswith("a"))]
                used.update(t.id for t in toks)
                per_col[c.id] = toks
            # a glued pair: the right half sits in its column, the left half belongs to the column just before it
            ids = [c.id for c in sec.columns]
            for cid, toks in list(per_col.items()):
                for t in toks:
                    if t.glued_from and t.id.endswith("b") and t.id[:-1] in lefts:
                        i = ids.index(cid)
                        if i > 0 and not per_col[ids[i - 1]]:
                            per_col[ids[i - 1]] = [lefts[t.id[:-1]]]
                            used.add(lefts[t.id[:-1]].id)
            mapped_ids = {t.id for c in sec.columns if c.candidates for t in per_col[c.id]}
            ocr_pdf = tk.format == "ocr-pdf" and tk.source_path
            stray = [t for t in cand if t.id not in used and t.kind == "amount"] + [
                t for t in row_toks if t.kind == "text" and re.fullmatch(r"\(?-?[\d.]+(,\d+)?-?\)?", t.text)
                and len(t.text) >= 4 and t.id not in mapped_ids]
            for c in sec.columns:
                cells[c.id] = _cell(c.id, per_col[c.id], teur, tk.number_style)
                cells[c.id].glued = any(t.glued_from for t in per_col[c.id])
                if stray and cells[c.id].status == "empty":  # an amount on this line found no column: zero is not safe
                    cells[c.id] = Cell(column=c.id, tokens=[], text="", cents=None, status="suspect_empty")
                elif ocr_pdf and cells[c.id].status == "empty" and c.candidates:
                    # OCR can drop a word: an empty cell is a zero only if the image shows no ink there either
                    from .ocr import ink_ratio
                    from .tokens import _page_from_frame
                    p = next(x for x in tk.pages if x["page"] == rb.page)
                    box = _page_from_frame([c.span[0], rb.band[0], c.anchor, rb.band[1]], p["direction"], p["width"], p["height"])
                    if ink_ratio(tk.source_path, rb.page, box) > 0.04:
                        cells[c.id] = Cell(column=c.id, tokens=[], text="", cents=None, status="suspect_empty")
            label_toks = [t for t in row_toks if t.id not in mapped_ids and t.kind not in ("amount", "dash")]
            sub_hint = any(re.search(r"\bunr\b|unr\.|unter", h, re.I) for h in sec.header_lines)
            ctx_codes = {w for c in rb.context for w in re.findall(r"\b\d{3,5}\b", c)}
            if len(label_toks) > 1 and label_toks[0].text in ctx_codes and re.fullmatch(r"\d+[a-z]?", label_toks[1].text or ""):
                label_toks = label_toks[1:]  # the account code printed in front of every asset line
            ident = _identity_from_label(label_toks, sub_hint) if rb.kind_guess == "asset" else {}
            pointer = {"page": rb.page, "bbox": _union([t.bbox for t in row_toks]) if row_toks else None,
                       "frame_band": rb.band}
        else:  # xlsx: the cell at (row, column)
            for c in sec.columns:
                t = by_cell.get((rb.sheet, rb.row, c.col))
                cells[c.id] = _cell(c.id, [t] if t else [], teur, tk.number_style)
            ident = {}
            if rb.kind_guess == "asset":
                for f, ci in sec.id_columns.items():
                    t = by_cell.get((rb.sheet, rb.row, ci))
                    if t:
                        ident[f] = t.text if t.kind != "number" else str(int(t.number)) if float(t.number).is_integer() else t.text
                if "acq_date" in ident:  # only a real date is a date
                    d = parse_date(ident["acq_date"])
                    ident["acq_date"] = d.isoformat() if d else None
                if "asset_no" not in ident:
                    label_toks = [by_cell[(rb.sheet, rb.row, ci)] for (sh, r, ci) in by_cell if sh == rb.sheet and r == rb.row
                                  and ci not in {c.col for c in sec.columns} and by_cell[(sh, r, ci)].kind in ("text", "int", "date")]
                    ident = {**_identity_from_label(sorted(label_toks, key=lambda t: t.col), False), **ident}
            pointer = {"sheet": rb.sheet, "cell": f"A{rb.row}", "row": rb.row}
        out.append(ExtractedRow(id=rb.id, section=rb.section, kind_guess=rb.kind_guess, label=rb.label, year=year,
                                identity=ident, cells=cells, pointer=pointer))
    return Extracted(doc_id=lay.doc_id, rows=out)
