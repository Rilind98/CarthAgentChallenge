"""Which blanks and errors come from the layout step, which from OCR, which from writing the Arbeitsversion.

    .venv/bin/python bench/attribute.py <run dir> [--truth out_key]

Per answer-key value (06_score.json compares the key-shaped output with <id>.key.json): every value that is not
correct gets one cause from the evidence the pipeline recorded (document status, tokens in the printed row, column
mapping, cell status and flags, printed text). The Arbeitsversion score (bench/score_av.py) is then split into what
the key comparison already explains and what only the workbook placement adds.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

OCR_FLAGS = ("low_ocr_conf", "ocr_law_fail", "suspect_empty", "unparsed", "ambiguous")
LAYOUT_FLAGS = ("field_uncertain", "mapping_unverified")


def _overlap(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    return ix * iy


def attribute_doc(run: Path, doc: str, key: dict) -> Counter:
    c = Counter()
    d = run / doc
    summ = {}
    if not (d / "04_doc.json").exists():
        for kr in key["rows"]:
            c[("blank", "document crashed", "layout")] += len({**kr.get("fields", {}), **(kr.get("net_fields") or {})})
        return c
    doc_json = json.load(open(d / "04_doc.json"))
    tk = json.load(open(d / "00_tokens.json"))
    ocr = tk["format"] in ("ocr-pdf", "image-pdf")
    abstain = doc_json["meta"].get("abstain")
    ours = doc_json["rows"]
    by_page = defaultdict(list)
    for t in tk["tokens"]:
        if t.get("page"):
            by_page[t["page"]].append(t)
    for kr in key["rows"]:
        truth = {**kr.get("fields", {}), **(kr.get("net_fields") or {})}
        kp = kr["pointer"]
        if abstain:
            src = "ocr" if "OCR" in abstain or "text layer" in abstain else "layout"
            c[("blank", f"document not read: {abstain[:40]}", src)] += len(truth)
            continue
        cands = [o for o in ours if o["pointer"].get("page") == kp.get("page") and o["pointer"].get("bbox")
                 and _overlap(o["pointer"]["bbox"], kp["bbox"]) > 0.5 * max(1e-6, (o["pointer"]["bbox"][2] - o["pointer"]["bbox"][0]) * (o["pointer"]["bbox"][3] - o["pointer"]["bbox"][1]))]
        o = cands[0] if cands else None
        if o is None:
            amounts = [t for t in by_page.get(kp.get("page"), []) if t["kind"] == "amount" and _overlap(t["bbox"], kp["bbox"]) > 0]
            src = "layout" if amounts else ("ocr" if ocr else "layout")
            c[("blank", "row not found" + ("" if amounts else " (no amount read on the printed line)"), src)] += len(truth)
            continue
        for f, v in truth.items():
            got = o["fields"].get(f, "missing")
            flag = (o["flags"].get(f) or "")
            if got == v:
                c[("correct", "", "")] += 1
            elif got == "missing":
                c[("blank", "field not mapped / column not found", "layout")] += 1
            elif got is None:
                if flag.startswith(OCR_FLAGS):
                    c[("blank", "flag " + flag.split(":")[0], "ocr" if ocr else "layout")] += 1
                elif flag.startswith(LAYOUT_FLAGS):
                    c[("blank", "flag " + flag.split(":")[0], "layout")] += 1
                else:
                    c[("blank", "flag " + (flag.split(":")[0] or "other"), "verification (laws)")] += 1
            else:
                ours_printed = o["printed"].get(f, "")
                key_printed = (kr.get("printed") or {}).get(f, "")
                if ours_printed == key_printed:
                    c[("incorrect", "right text, wrong field/row", "layout")] += 1
                elif not ours_printed and key_printed:
                    c[("incorrect", "printed value missed (read as empty)", "ocr" if ocr else "layout")] += 1
                else:
                    c[("incorrect", "characters misread", "ocr" if ocr else "layout")] += 1
        for f, v in o["fields"].items():
            if f not in truth and v is not None:
                c[("incorrect", "value in a field the key does not have", "layout")] += 1
    matched = set()
    return c


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("run", type=Path)
    ap.add_argument("--truth", type=Path, default=Path("out_key"))
    ap.add_argument("--corpus", type=Path, default=Path("bench/corpus_pdf"))
    a = ap.parse_args()
    total = Counter()
    for kp in sorted(a.corpus.glob("*.key.json")):
        doc = kp.name[:-9]
        total += attribute_doc(a.run, doc, json.load(open(kp)))
    by_src = defaultdict(Counter)
    for (outcome, why, src), n in total.items():
        if outcome != "correct":
            by_src[outcome][src] += n
    print("== per answer-key value (key-shaped JSON output)")
    print(f"correct {total[('correct', '', '')]}")
    for outcome in ("blank", "incorrect"):
        print(f"{outcome}: {sum(by_src[outcome].values())}  by source: {dict(by_src[outcome])}")
        for (o, why, src), n in sorted(total.items(), key=lambda kv: -kv[1]):
            if o == outcome:
                print(f"    {n:6}  {src:22} {why}")
    # workbook: what the placement adds on top
    sys.path.insert(0, str(Path(__file__).parent))
    import score_av
    av = Counter()
    for tp in sorted(a.truth.glob("Arbeitsversion_*.xlsx")):
        doc = tp.stem.removeprefix("Arbeitsversion_")
        if (a.corpus / f"{doc}.pdf").exists():
            av += score_av.score_doc(tp, a.run / doc / f"Arbeitsversion_{doc}.xlsx")
    key_correct = total[("correct", "", "")]
    print("\n== Arbeitsversion (workbook) against out_key")
    print(f"truth values {av['truth_values']}, correct {av['correct']}, blank {av['blank']}, "
          f"incorrect {av['incorrect'] + av['incorrect_extra'] + av['incorrect_spurious']} "
          f"(different number {av['incorrect']}, number where truth has none {av['incorrect_extra']}, "
          f"row not in truth {av['incorrect_spurious']})")


if __name__ == "__main__":
    main()
