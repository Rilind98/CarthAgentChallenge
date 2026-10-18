"""Score agent output against truth/ (made by scripts/keys_to_truth.py), with the brief's points: +1 / -3.

A truth value is matched by: same document and year, same column, same page, and the centre of the output's box
inside the truth row box (PDF truth boxes are whole-row boxes, +2 pt). Each truth value is used once.
  correct   value equal                                  +1
  wrong     value differs                                -3
  extra     output with no truth value at that place     -3 (the brief: a value whose pointer does not show it)
  missing   truth value the output does not have          0 here (penalty unknown until we have the scorer)

Usage: python3 -I agent/score.py OUTPUT_DIR TRUTH_DIR [--only SUBSTR] [--show N]
"""
import argparse
import json
from collections import Counter
from pathlib import Path


def score_file(out: list[dict], truth: list[dict]) -> tuple[Counter, list]:
    used, c, bad = set(), Counter(), []
    for o in out:
        cx, cy = (o["bbox"][0] + o["bbox"][2]) / 2, (o["bbox"][1] + o["bbox"][3]) / 2
        hit = None
        for j, t in enumerate(truth):
            if j in used or t["column"] != o["column"] or t.get("page") != o["page"]:
                continue
            x0, y0, x1, y1 = t["bbox"]
            if x0 - 2 <= cx <= x1 + 2 and y0 - 2 <= cy <= y1 + 2:
                if t["value"] == o["value"]:
                    hit = j
                    break
                hit = hit if hit is not None else j
        if hit is None:
            c["extra"] += 1
            bad.append(("extra", o, None))
            continue
        used.add(hit)
        if truth[hit]["value"] == o["value"]:
            c["correct"] += 1
        else:
            c["wrong"] += 1
            bad.append(("wrong", o, truth[hit]))
    c["missing"] = len(truth) - len(used)
    c["truth"] = len(truth)
    return c, bad


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("output", type=Path)
    ap.add_argument("truth", type=Path)
    ap.add_argument("--only", default="")
    ap.add_argument("--show", type=int, default=0)
    a = ap.parse_args()
    tot = Counter()
    for of in sorted(a.output.glob("*_[12][0-9][0-9][0-9].json")):
        if a.only not in of.name:
            continue
        tf = a.truth / of.name
        if not tf.exists():
            print(f"{of.name:40s} no truth file (wrong year?)")
            continue
        c, bad = score_file(json.loads(of.read_text()), json.loads(tf.read_text()))
        tot.update(c)
        pts = c["correct"] - 3 * (c["wrong"] + c["extra"])
        print(f"{of.stem:34s} correct {c['correct']:5d}  wrong {c['wrong']:4d}  extra {c['extra']:4d}  "
              f"missing {c['missing']:4d}/{c['truth']:<5d} points {pts:6d}")
        for kind, o, t in bad[: a.show]:
            print(f"     {kind:5s} p{o['page']} {o['column']:13s} out {o['value']:>12} "
                  f"{'truth ' + str(t['value']) if t else ''}  row {o['row'][:40]}")
    pts = tot["correct"] - 3 * (tot["wrong"] + tot["extra"])
    print(f"\nTOTAL correct {tot['correct']}  wrong {tot['wrong']}  extra {tot['extra']}  "
          f"missing {tot['missing']}/{tot['truth']}  points {pts}  "
          f"(max {tot['truth']}, precision {tot['correct'] / max(1, tot['correct'] + tot['wrong'] + tot['extra']):.2%})")


if __name__ == "__main__":
    main()
