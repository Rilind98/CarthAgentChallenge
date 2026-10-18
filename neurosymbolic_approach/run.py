"""Run the pipeline on a folder of hand-over documents and write every stage's output.

    .venv/bin/python run.py <input folder> <output folder> [--only f01_net_spiegel ...]

For each document <id> it writes <output>/<id>/:
    00_tokens.json     step 0  every printed string with its box (text layer / xlsx cells)
    01_layout_geometry.json / 01_vlm.json  step 1  code's geometry, then the VLM's labels (prompt input + answer)
    01_layout.json     step 1  row and column boxes, header text, field candidates (no figures)
    02_extracted.json  step 2  each row x column cell copied and parsed to cents by code
    03_checks.json     step 3  chosen column mapping, law checks, totals, table end, completeness
    03_detections.json step 3b noise / irregularities found, named like the keys' `planted` types
    04_doc.json        step 4  key-shaped document: one row per printed row, fields + pointers + flags
    05_values.json     step 4  one record per value (the hand-in JSON)
    Arbeitsversion_<id>.xlsx  step 4b  the filled template (needs --template, default ../Arbeitsversion-merged-layout-20.xlsx)
    06_score.json      dev     comparison with <id>.key.json when it exists
and <output>/summary.md + summary.json over all documents.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import traceback
from collections import defaultdict
from pathlib import Path

from pipeline import arbeitsversion, vlm
from pipeline.detect import detect, score_detection
from pipeline.extract import extract
from pipeline.layout import propose
from pipeline.output import build
from pipeline.score import score
from pipeline.tokens import read_tokens
from pipeline.verify import verify


def _dump(path: Path, obj) -> None:
    if hasattr(obj, "model_dump"):
        obj = obj.model_dump()
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=1, default=str) + "\n", encoding="utf-8")


def documents(root: Path, only: list[str]) -> list[Path]:
    out = []
    for p in sorted(root.rglob("*")):
        rel = p.relative_to(root).parts
        if any(part.startswith(("_", ".")) or part in ("OUT", "out") for part in rel[:-1]):
            continue
        if p.suffix.lower() in (".pdf", ".xlsx", ".xlsm") and (not only or any(o in p.name for o in only)):
            out.append(p)
    return out


TEMPLATE = Path(__file__).resolve().parent.parent / "Arbeitsversion-merged-layout-20.xlsx"


def run_one(path: Path, out_root: Path, template: Path | None = TEMPLATE) -> dict:
    doc_id = path.stem
    od = out_root / doc_id
    if od.exists():  # a document folder holds only this run's outputs
        for f in od.iterdir():
            f.unlink()
    od.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    tk = read_tokens(path, doc_id)
    _dump(od / "00_tokens.json", tk)
    lay = propose(tk)
    if vlm.available() and lay.sections and not lay.abstain:  # the VLM labels the geometry code found
        _dump(od / "01_layout_geometry.json", lay)
        lay, vlog = vlm.label_layout(path, lay, tk.pages)
        _dump(od / "01_vlm.json", vlog)
    _dump(od / "01_layout.json", lay)
    ex = extract(tk, lay)
    _dump(od / "02_extracted.json", ex)
    report, decisions = verify(tk, lay, ex)
    _dump(od / "03_checks.json", report)
    det = detect(tk, lay, ex, report, decisions)
    _dump(od / "03_detections.json", det)
    meta = {"doc_id": doc_id, "source": str(path), "format": tk.format}
    doc, values = build(tk, lay, ex, report, decisions, meta)
    doc["detections"] = det["findings"]
    doc["meta"]["variant_guess"] = det["variant_guess"]
    _dump(od / "04_doc.json", doc)
    _dump(od / "05_values.json", values)
    av_info = None
    if template and template.exists():
        av_info = arbeitsversion.write(template, od / f"Arbeitsversion_{doc_id}.xlsx", doc, det["findings"])
    res = {"doc_id": doc_id, "family": path.parent.name, "format": tk.format, "mode": lay.mode, "abstain": lay.abstain,
           "rows": len(doc["rows"]), "values": sum(1 for v in values if v.get("value") is not None),
           "blank": sum(1 for v in values if v.get("field") and v.get("value") is None),
           "gates": [s["gate"] for s in report["sections"]], "arbeitsversion": av_info,
           "seconds": round(time.time() - t0, 2)}
    key_path = path.with_name(doc_id + ".key.json")
    if key_path.exists():
        key = json.loads(key_path.read_text(encoding="utf-8"))
        sc = score(doc, key)
        sc["detection"] = score_detection(det, key, readable=tk.format != "image-pdf" and not lay.abstain)
        _dump(od / "06_score.json", sc)
        res["detection"] = sc["detection"]
        res["key_format"] = key["meta"]["format"]
        res.update({k: sc[k] for k in ("values_in_key", "correct", "wrong", "blank", "missed", "extra", "spurious", "score",
                                        "kind_accuracy", "asset_id_accuracy")})
    return res


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--only", nargs="*", default=[])
    ap.add_argument("--template", type=Path, default=TEMPLATE, help="empty Arbeitsversion to fill")
    a = ap.parse_args()
    a.output.mkdir(parents=True, exist_ok=True)
    results = []
    for p in documents(a.input, a.only):
        try:
            r = run_one(p, a.output, a.template)
        except Exception as e:  # one broken document must not stop the run
            r = {"doc_id": p.stem, "family": p.parent.name, "error": f"{type(e).__name__}: {e}"}
            (a.output / p.stem).mkdir(parents=True, exist_ok=True)
            (a.output / p.stem / "error.txt").write_text(traceback.format_exc())
        results.append(r)
        line = f"{r['doc_id']:28} {r.get('format', '?'):9} {r.get('mode', '?'):5} rows={r.get('rows', 0):4} "
        if "score" in r:
            line += f"key={r['values_in_key']:5} ok={r['correct']:5} wrong={r['wrong'] + r['extra'] + r['spurious']:4} blank={r['blank'] + r['missed']:5} score={r['score']:6}"
        print(line + (f"  ERROR {r['error']}" if "error" in r else f"  {r.get('abstain') or ''}"), flush=True)
    _dump(a.output / "summary.json", results)
    fam = defaultdict(lambda: defaultdict(int))
    for r in results:
        for k in ("values_in_key", "correct", "wrong", "extra", "spurious", "blank", "missed", "score"):
            fam[r["family"]][k] += r.get(k, 0) or 0
        fam[r["family"]]["docs"] += 1
        fam[r["family"]]["abstained"] += bool(r.get("abstain"))
        fam[r["family"]]["errors"] += "error" in r
    lines = ["# Pipeline run", "", "| family | docs | abstained | values in key | correct | wrong | blank/missed | score |",
             "|---|---|---|---|---|---|---|---|"]
    tot = defaultdict(int)
    for f, s in sorted(fam.items()):
        w = s["wrong"] + s["extra"] + s["spurious"]
        lines.append(f"| {f} | {s['docs']} | {s['abstained']} | {s['values_in_key']} | {s['correct']} | {w} | {s['blank'] + s['missed']} | {s['score']} |")
        for k, v in s.items():
            tot[k] += v
    w = tot["wrong"] + tot["extra"] + tot["spurious"]
    lines.append(f"| **all** | {tot['docs']} | {tot['abstained']} | {tot['values_in_key']} | {tot['correct']} | {w} | {tot['blank'] + tot['missed']} | **{tot['score']}** |")
    lines += ["", "| doc | format | mode | rows | correct | wrong | blank/missed | score | note |", "|---|---|---|---|---|---|---|---|---|"]
    for r in results:
        if "score" in r:
            lines.append(f"| {r['doc_id']} | {r['format']} | {r['mode']} | {r['rows']} | {r['correct']} | {r['wrong'] + r['extra'] + r['spurious']} | "
                         f"{r['blank'] + r['missed']} | {r['score']} | {r.get('abstain') or r.get('error') or ''} |")
    det = defaultdict(lambda: defaultdict(int))
    for r in results:
        d = r.get("detection")
        if not d or not d["readable"]:
            continue
        for t in d["hit"]:
            det[t]["hit"] += 1
        for t in d["missed"]:
            det[t]["missed"] += 1
        for t in d["false_alarm"]:
            det[t]["false_alarm"] += 1
    lines += ["", "## Detection of planted problems (documents with a text layer that were read)", "",
              "| type | hit | missed | false alarm |", "|---|---|---|---|"]
    for t, c in sorted(det.items()):
        lines.append(f"| {t} | {c['hit']} | {c['missed']} | {c['false_alarm']} |")
    (a.output / "summary.md").write_text("\n".join(lines) + "\n")
    print(f"\nall: values in key {tot['values_in_key']}, correct {tot['correct']}, wrong {w}, blank/missed {tot['blank'] + tot['missed']}, score {tot['score']}")


if __name__ == "__main__":
    sys.exit(main())
