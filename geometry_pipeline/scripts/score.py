"""Score extraction files against truth/, the way the brief counts: +1 correct, 0 left blank, -3 wrong.

Records are matched by where they point, not by row label (labels of sum rows are ambiguous):
  PDF  - same page and column; the truth line box nearest to the output box centre (0 when inside, ties by
         centre distance), at most 12 pt away (scans are slightly skewed against the clean layout)
  xlsx - same sheet and column, and same spreadsheet row
A record with "flag" counts as left blank. An output record that matches no truth value counts as wrong.

Usage: .venv/bin/python scripts/score.py OUT_DIR [TRUTH_DIR]
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path


def centre(b: list) -> tuple[float, float]:
    return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2


def match(o: dict, truth: list[dict], used: set) -> int | None:
    cands = [i for i, t in enumerate(truth) if i not in used and t["column"] == o["column"]]
    if "sheet" in o:
        row = lambda c: re.sub(r"\D", "", c)  # noqa: E731
        hit = [i for i in cands if truth[i].get("sheet") == o["sheet"] and row(truth[i]["cell"]) == row(o["cell"])]
        return hit[0] if hit else None
    cands = [i for i in cands if truth[i].get("page") == o.get("page")]
    if not cands or "bbox" not in o:
        return None
    x, y = centre(o["bbox"])

    def dist(i: int) -> tuple[float, float]:  # distance to the box (0 inside), then to its centre
        x0, y0, x1, y1 = truth[i]["bbox"]
        cx, cy = centre(truth[i]["bbox"])
        return max(x0 - x, 0, x - x1) + max(y0 - y, 0, y - y1), abs(x - cx) + abs(y - cy)

    best = min(cands, key=dist)
    return best if dist(best)[0] <= 12 else None


def score(out: list[dict], truth: list[dict]) -> Counter:
    c, used = Counter(), set()
    for o in out:
        i = match(o, truth, used)
        if o.get("flag"):
            c["flagged"] += 1
            if i is not None:
                used.add(i)
        elif i is None:
            c["unmatched"] += 1
        else:
            used.add(i)
            c["correct" if o["value"] == truth[i]["value"] else "wrong"] += 1
    c["missed"] = len(truth) - len(used)
    c["truth"] = len(truth)
    return c


def main(out_dir: Path, truth_dir: Path) -> None:
    tot = Counter()
    for f in sorted(out_dir.glob("*_[0-9][0-9][0-9][0-9].json")):
        t = truth_dir / f.name
        if not t.exists():
            print(f"{f.name:36s} no truth file (wrong year or doc id?)")
            continue
        c = score(json.loads(f.read_text()), json.loads(t.read_text()))
        tot.update(c)
        pts = c["correct"] - 3 * (c["wrong"] + c["unmatched"])
        print(f"{f.name:36s} correct {c['correct']:4d}  wrong {c['wrong']:3d}  unmatched {c['unmatched']:3d}  "
              f"flagged {c['flagged']:3d}  missed {c['missed']:4d}/{c['truth']:<4d}  points {pts}")
    pts = tot["correct"] - 3 * (tot["wrong"] + tot["unmatched"])
    print(f"\nTOTAL correct {tot['correct']}  wrong {tot['wrong']}  unmatched {tot['unmatched']}  flagged {tot['flagged']}  "
          f"missed {tot['missed']}/{tot['truth']}  points {pts}  (max {tot['truth']})")


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent.parent / "truth")
