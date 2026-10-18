"""Convert the corpus answer keys into the internal working format.

One file per document and year, <doc_id>_<year>.json, holding a flat list with one record per value:
    {"row": str, "column": str, "value": int (cents), "page": int, "bbox": [x0, y0, x1, y1]}   (PDF)
    {"row": str, "column": str, "value": int (cents), "sheet": str, "cell": str}               (xlsx)

PDF keys only carry a box per line, so `bbox` here is the row's line box, not the box of the single value.
Sum rows carry no printed label in the keys, so their `row` is "<kind> <group>" (e.g. "class_subtotal 0027").

Usage: python3 -I scripts/keys_to_truth.py CORPUS_DIR OUT_DIR
"""
import json
import sys
from collections import defaultdict
from pathlib import Path


def row_label(r: dict) -> str:
    if r["kind"] == "asset":
        return " ".join(s for s in (r.get("asset_no"), r.get("sub_no"), r.get("name")) if s)
    group = r.get("class") or r.get("account") or r.get("position") or ""
    return f"{r['kind']} {group}".strip()


def records(r: dict) -> list[dict]:
    p, row, out = r["pointer"], row_label(r), []
    for column, value in {**r["fields"], **r.get("net_fields", {})}.items():
        rec = {"row": row, "column": column, "value": value}
        if "page" in p:
            rec |= {"page": p["page"], "bbox": p["bbox"]}
        else:
            rec |= {"sheet": p["sheet"], "cell": p["cells"].get(column, p["cell"])}
        out.append(rec)
    return out


def main(corpus: Path, out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    n_files = n_values = 0
    for key in sorted(corpus.glob("f*/*.key.json")):
        doc_id = key.name.removesuffix(".key.json")
        by_year = defaultdict(list)
        for r in json.loads(key.read_text())["rows"]:
            by_year[r["year"]].extend(records(r))
        for year, recs in by_year.items():
            text = "[\n" + ",\n".join("  " + json.dumps(x, ensure_ascii=False) for x in recs) + "\n]\n"
            (out / f"{doc_id}_{year}.json").write_text(text)
            n_files += 1
            n_values += len(recs)
    print(f"{n_files} files, {n_values} values -> {out}")


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]))
