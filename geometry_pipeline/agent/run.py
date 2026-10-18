"""Run the agent on a folder: one <doc_id>_<year>.json per document and year, flat list of values.

    record: {"row", "column", "value" (int cents), "page", "bbox"}

Usage: .venv/bin/python agent/run.py INPUT_DIR OUTPUT_DIR [--only SUBSTR]
"""
import argparse
import json
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from extract import pdf_pages  # noqa: E402
from xlsx import xlsx_pages  # noqa: E402
from grid import build  # noqa: E402
from layout import decide, group_pages, kind_of, load_env  # noqa: E402
from ocr import recheck_page, review_page  # noqa: E402
from records import page_records, sum_suspects  # noqa: E402


def run_doc(pdf: Path, out_dir: Path) -> dict:
    pages = xlsx_pages(pdf) if pdf.suffix == ".xlsx" else pdf_pages(pdf)
    if not any(p["words"] for p in pages):
        return {"doc": pdf.stem, "skipped": "no words on any page"}
    images = [p["pdf_page"].get_pixmap(dpi=100).tobytes("png") for p in pages[:3] if "pdf_page" in p]

    def grids_and_layout(teur: bool):
        first = [build(p["page"], p["words"], teur) for p in pages]
        groups = group_pages(first, separate=pdf.suffix == ".xlsx")
        strips_of = {pg: gr["strips"] for gr in groups for pg in gr["pages"]}
        grids = [build(p["page"], p["words"], teur, strips_of.get(p["page"])) for p in pages]
        return grids, groups, decide(grids, groups, images)

    grids, groups, lay = grids_and_layout(False)
    teur = lay.get("unit") == "TEUR"
    if teur:  # rebuild: in TEUR amounts have no decimals
        grids, groups, lay = grids_and_layout(True)
        lay["unit"] = "TEUR"
    page_info = {p["page"]: p for p in lay.get("pages", [])}
    group_of = {pg: gi for gi, gr in enumerate(groups) for pg in gr["pages"]}
    rules_by_group = {g["group"]: g.get("rules", []) for g in lay.get("groups", [])}
    by_year = defaultdict(list)
    review = [p for p in pages if "scan" in p and page_info.get(p["page"], {}).get("role") == "table"
              and p["page"] in group_of]
    if review:  # scanned table pages: proofread tesseract's words now that the layout says what the strips are
        strips_of = {pg: gr["strips"] for gr in groups for pg in gr["pages"]}

        def layout_of(p):
            gi = group_of[p["page"]]
            return {"page": p["page"], "teur": teur, "strips": strips_of[p["page"]],
                    "rules": rules_by_group.get(gi, []), "kinds": lay.get("kinds", []),
                    "headers": groups[gi].get("headers", {})}

        def proofread(p):
            return review_page(p["scan"], layout_of(p))
        with ThreadPoolExecutor(len(review)) as ex:
            for p, (words, usage) in zip(review, ex.map(proofread, review)):
                p.update(words=words, ocr_usage=usage)
        grids = [build(p["page"], p["words"], teur, strips_of.get(p["page"])) for p in pages]
        # column sums over the whole document: re-read the cells one misread digit away from a broken sum
        recs = [(row, fields, g.page) for g in grids
                if page_info.get(g.page, {}).get("role") == "table" and g.page in group_of
                for row, fields in page_records(g, rules_by_group.get(group_of[g.page], []), lay.get("kinds", []),
                                                groups[group_of[g.page]].get("headers", {}), teur)]
        suspects = sum_suspects(recs, lay.get("kinds", []), teur)

        def recheck(p):
            return recheck_page(p["scan"], layout_of(p), [s for s in suspects if s["page"] == p["page"]])
        with ThreadPoolExecutor(len(review)) as ex:
            for p, (words, usage) in zip(review, ex.map(recheck, review)):
                p["words"] = words
                p["ocr_usage"] = {"cost_usd": p["ocr_usage"]["cost_usd"] + usage["cost_usd"]}
        grids = [build(p["page"], p["words"], teur, strips_of.get(p["page"])) for p in pages]
    for p, g in zip(pages, grids):
        info = page_info.get(g.page, {})
        if info.get("role") != "table" or g.page not in group_of:
            continue
        rules = rules_by_group.get(group_of[g.page], [])
        headers = groups[group_of[g.page]].get("headers", {})
        for row, fields in page_records(g, rules, lay.get("kinds", []), headers, teur):
            if kind_of(row, lay.get("kinds", [])) == "ignore":  # not a schedule row (headers, reconciliation)
                continue
            top = min(box[1] for _, box, _ in fields.values())  # the record's first line
            for field, (v, box, printed) in fields.items():
                if "sheet" in p:  # xlsx: a zero from an empty cell points at column A of the record's first row
                    cell = p["to_cell"](box if printed else (0, top, 1, top + 1))
                    where = {"sheet": p["sheet"], "cell": cell if printed else "A" + cell.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")}
                else:
                    where = {"page": g.page, "bbox": p["to_page"](box)}
                by_year[info.get("year")].append({"row": row, "kind": kind_of(row, lay.get("kinds", [])),
                                                  "column": field, "value": v, **where})
    doc_id = pdf.stem
    for year, recs in by_year.items():
        text = "[\n" + ",\n".join("  " + json.dumps(x, ensure_ascii=False) for x in recs) + "\n]\n"
        (out_dir / f"{doc_id}_{year}.json").write_text(text)
    lay["_groups"] = [{k: v for k, v in gr.items()} for gr in groups]
    (out_dir / f"{doc_id}.layout.json").write_text(json.dumps(lay, ensure_ascii=False, indent=1, default=str))
    ocr = [p["ocr_usage"] for p in pages if "ocr_usage" in p]
    return {"doc": doc_id, "values": sum(map(len, by_year.values())), "usage": lay.get("_usage"),
            **({"ocr_pages": len(ocr), "ocr_cost_usd": round(sum(u["cost_usd"] for u in ocr), 2)} if ocr else {})}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--only", default="")
    ap.add_argument("--jobs", type=int, default=6)
    a = ap.parse_args()
    load_env(Path(__file__).parent.parent / ".env")
    a.output.mkdir(parents=True, exist_ok=True)
    docs = [p for p in sorted(a.input.rglob("*")) if p.suffix in (".pdf", ".xlsx") and a.only in p.name
            and not p.name.startswith((".", "~$")) and "_generator" not in p.parts]

    def safe(p):
        try:
            return run_doc(p, a.output)
        except Exception as e:  # keep going; report
            return {"doc": p.stem, "error": f"{type(e).__name__}: {e}"}

    with ThreadPoolExecutor(a.jobs) as ex:
        for r in ex.map(safe, docs):
            print(r, flush=True)


if __name__ == "__main__":
    main()
