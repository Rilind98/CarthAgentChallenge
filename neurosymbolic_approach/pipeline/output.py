"""Step 4 - output. Key-shaped document JSON (one row per printed row) and one record per value, both pydantic-validated.

Every value carries its pointer: PDF page + box of the printed characters (or of the empty column box when the
cell is blank, which counts as zero), xlsx sheet + cell.
"""

from __future__ import annotations

from .schema import Extracted, Layout, OutRow, Tokens, ValueRecord
from .tokens import _page_from_frame


def build(tk: Tokens, lay: Layout, ex: Extracted, report: dict, decisions: dict, meta: dict) -> tuple[dict, list[dict]]:
    pages = {p["page"]: p for p in tk.pages}
    secs = {s.id: s for s in lay.sections}
    mappings = {s["id"]: s["mapping"] for s in report["sections"]}
    contexts = {r.id: r.context for r in lay.rows}
    rows, values = [], []
    for r in ex.rows:
        d = decisions.get(r.id)
        if not d or not d["emit"]:
            continue
        mapping = mappings[r.section]
        field_ptr = {}
        for cid, f in mapping.items():
            if not f:
                continue
            c = r.cells[cid]
            if r.pointer.get("sheet"):
                col = next(x.col for x in secs[r.section].columns if x.id == cid)
                from openpyxl.utils import get_column_letter
                field_ptr[f] = {"sheet": r.pointer["sheet"], "cell": c.cell or f"{get_column_letter(col)}{r.pointer['row']}"}
            else:
                box = c.box
                if box is None:  # empty cell: the column box on this row
                    col = next(x for x in secs[r.section].columns if x.id == cid)
                    p = pages[r.pointer["page"]]
                    fr = [col.span[0], r.pointer["frame_band"][0], col.anchor, r.pointer["frame_band"][1]]
                    box = [round(v, 2) for v in _page_from_frame(fr, p["direction"], p["width"], p["height"])]
                field_ptr[f] = {"page": r.pointer["page"], "bbox": box}
        ident = r.identity or {}
        pointer = dict(r.pointer)
        pointer["fields"] = field_ptr
        out = OutRow(id=r.id, section=r.section, sums=d.get("sums") or [], kind=d["kind"], level=d.get("level", 0), year=r.year,
                     asset_no=ident.get("asset_no"), sub_no=ident.get("sub_no"), name=ident.get("name"),
                     acq_date=ident.get("acq_date"), method=ident.get("method"), label=r.label,
                     context=contexts.get(r.id, []), fields=d["fields"], printed=d["printed"], pointer=pointer,
                     flags=d["flags"])
        rows.append(out.model_dump())
        for f, v in d["fields"].items():
            values.append(ValueRecord(doc_id=lay.doc_id, row_id=r.id, kind=d["kind"], year=r.year,
                                      asset_no=out.asset_no, sub_no=out.sub_no, label=r.label, field=f, value=v,
                                      printed=d["printed"].get(f), pointer=field_ptr.get(f, {}),
                                      flag=d["flags"].get(f) or d["flags"].get("*"),
                                      conf=next((r.cells[cid].conf for cid, g in mapping.items() if g == f), None)).model_dump())
    doc = {"meta": {**meta, "proposer": lay.proposer, "mode": lay.mode, "abstain": lay.abstain,
                    "sections": [{"id": s.id, "period_end": s.period_end, "unit": s.unit, "pages": s.pages,
                                  "sheet": s.sheet, "leaf_level": s.leaf_level} for s in lay.sections]},
           "rows": rows}
    if lay.abstain:
        values.append({"doc_id": lay.doc_id, "row_id": None, "field": None, "value": None,
                       "flag": "document_blank", "reason": lay.abstain})
    return doc, values
