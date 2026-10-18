"""Command-line tool of the OCR proofreading agent (ocr.review_page): the only command its Bash may run (the
session allows `./tool *` and reading files in its directory, nothing else). Claude supplies the eyes; the
arithmetic and every change to the words go through the fixed code here.

State of one page in DIR: page.png (upright, deskewed), words.json [{id, box (px), text}], layout.json (strips,
rules, kinds, headers of the page's group, teur, dpi), edits.json (every change, for the audit).

Usage: python agent/ocr_tools.py DIR view [x0 y0 x1 y1] [--boxes] | check | fix ID TEXT | merge ID ID... [--text T]
       | add x0 y0 x1 y1 TEXT
"""
import argparse
import json
import sys
from pathlib import Path

import cv2

sys.path.insert(0, str(Path(__file__).parent))

from grid import Word, build  # noqa: E402
import records  # noqa: E402
from records import SIGNED, broken_laws, digit_distance, page_records  # noqa: E402

DIR = Path(sys.argv[1])
PAGE = cv2.imread(str(DIR / "page.png"), cv2.IMREAD_GRAYSCALE)
LAYOUT = json.loads((DIR / "layout.json").read_text())
S = 72 / LAYOUT["dpi"]  # px -> pt

def load() -> list[dict]:
    return json.loads((DIR / "words.json").read_text())


def save(words: list[dict], edit: dict) -> None:
    (DIR / "words.json").write_text(json.dumps(words, ensure_ascii=False))
    path = DIR / "edits.json"
    edits = json.loads(path.read_text()) if path.exists() else []
    path.write_text(json.dumps(edits + [edit], ensure_ascii=False))


def fmt(cents: int) -> str:
    return records.fmt(cents, LAYOUT["teur"])


def ids_in(words: list[dict], box: tuple) -> list[int]:
    """Words (px) that overlap a cell box (pt)."""
    x0, y0, x1, y1 = (c / S for c in box)
    return [w["id"] for w in words if w["box"][0] < x1 and w["box"][2] > x0 and w["box"][1] < y1 and w["box"][3] > y0]


def view(x0: int = 0, y0: int = 0, x1: int | None = None, y1: int | None = None, boxes: bool = False) -> str:
    """Writes the image of a page region (pixels; default the whole page), enlarged up to 2x, to view_<n>.png and
    lists the words whose centre lies in it as `id [x0,y0,x1,y1] text`, one text line per row. boxes=True draws
    the word boxes."""
    h, w = PAGE.shape
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(w, x1 if x1 is not None else w), min(h, y1 if y1 is not None else h)
    if x1 <= x0 or y1 <= y0:
        return "empty region"
    words = [v for v in load() if x0 <= (v["box"][0] + v["box"][2]) / 2 <= x1
             and y0 <= (v["box"][1] + v["box"][3]) / 2 <= y1]
    crop = cv2.cvtColor(PAGE[y0:y1, x0:x1], cv2.COLOR_GRAY2BGR)
    if boxes:
        for v in words:
            a, b, c, d = (int(round(t)) for t in v["box"])
            cv2.rectangle(crop, (a - x0, b - y0), (c - x0, d - y0), (0, 0, 255), 1)
    s = min(2.0, 1600 / crop.shape[1], 1600 / crop.shape[0])
    crop = cv2.resize(crop, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC if s > 1 else cv2.INTER_AREA)
    lines, last = [], None
    for v in sorted(words, key=lambda v: (v["box"][1] + v["box"][3]) / 2):
        yc = (v["box"][1] + v["box"][3]) / 2
        if last is None or yc - last > 15:
            lines.append([])
        lines[-1].append(v)
        last = yc
    listing = "\n".join("  ".join(f"{v['id']} [{','.join(str(round(c)) for c in v['box'])}] {v['text']}"
                                  for v in sorted(ln, key=lambda v: v["box"][0])) for ln in lines)
    path = DIR / f"view_{len(list(DIR.glob('view_*.png')))}.png"
    cv2.imwrite(str(path), crop)
    return (f"image: {path.name} (read it), region [{x0},{y0},{x1},{y1}] at {s:.2f}x\n"
            f"words:\n{listing or '(none)'}")


def check() -> str:
    """The row arithmetic on the current words, computed in code: rows that do not add up, their cells with word
    ids, and per cell the value that alone would make the row add up. Also numbers the code cannot read."""
    words = load()
    grid = build(LAYOUT["page"], [Word(*(c * S for c in v["box"]), v["text"]) for v in words if v["text"].strip()],
                 LAYOUT["teur"], [tuple(s) for s in LAYOUT["strips"]])
    out = []
    for ln in grid.lines:
        for i in sorted(ln.unreadable):
            a, b = grid.strips[i]
            ids = ids_in(words, (a, ln.y0, b, ln.y1))
            out.append(f"unreadable: row '{ln.label[:50]}' column S{i}: words {ids} "
                       + " ".join(repr(v["text"]) for v in words if v["id"] in ids))
    recs = page_records(grid, LAYOUT["rules"], LAYOUT["kinds"], LAYOUT["headers"], LAYOUT["teur"])
    for row, fields in recs:
        for coefs, residual in broken_laws(fields):
            eq = " ".join(f"{'+' if c > 0 else '-'} {f}" for f, c in coefs.items()).lstrip("+ ")
            out.append(f"row '{row[:50]}': {eq} = 0 is off by {fmt(residual)}")
            for f, c in coefs.items():
                v, box, printed = fields[f]
                need = v - residual // c
                shown = fmt(v) if printed else "empty"
                if need < 0 and f not in SIGNED:
                    hint = "-"  # no value of this cell alone can fix the row
                else:
                    hint = f"would need {fmt(need)}"
                    if printed and v and need and digit_distance(fmt(v), fmt(need)) <= 1:
                        hint += "  <- one digit apart"
                out.append(f"    {f:14s} {shown:>18s}  words {ids_in(words, box)}  {hint}")
    return f"{len(recs)} records on the page.\n" + ("\n".join(out) if out else "Every row adds up.")


def fix_word(id: int, text: str) -> str:
    """Set word `id` to the text printed there. Several space-separated tokens split its box between them by
    length (two numbers tesseract glued together); empty text deletes the word (noise)."""
    words = load()
    w = next((v for v in words if v["id"] == id), None)
    if w is None:
        return f"no word {id}"
    tokens = text.split()
    rest = [v for v in words if v["id"] != id]
    new = []
    if tokens:
        x0, y0, x1, y1 = w["box"]
        per, x, nid = (x1 - x0) / sum(map(len, tokens)), x0, max(v["id"] for v in words) + 1
        for j, t in enumerate(tokens):
            new.append({"id": id if j == 0 else nid + j - 1, "box": [x, y0, x + per * len(t), y1], "text": t})
            x += per * len(t)
    save(rest + new, {"op": "fix", "id": id, "old": w["text"], "new": text})
    return "deleted" if not new else "now " + "  ".join(f"{v['id']} {v['text']}" for v in new)


def merge_words(ids: list[int], text: str | None = None) -> str:
    """Join words that are one printed token (a number tesseract split): union box, text joined left to right
    unless `text` is given."""
    words = load()
    parts = sorted((v for v in words if v["id"] in ids), key=lambda v: v["box"][0])
    if len(parts) < 2:
        return "need at least two existing ids"
    box = [min(v["box"][0] for v in parts), min(v["box"][1] for v in parts),
           max(v["box"][2] for v in parts), max(v["box"][3] for v in parts)]
    merged = {"id": parts[0]["id"], "box": box, "text": text or "".join(v["text"] for v in parts)}
    save([v for v in words if v["id"] not in ids] + [merged],
         {"op": "merge", "ids": ids, "old": [v["text"] for v in parts], "new": merged["text"]})
    return f"now {merged['id']} {merged['text']}"


def add_word(x0: int, y0: int, x1: int, y1: int, text: str) -> str:
    """Add a token tesseract missed, with the pixel box of its ink."""
    words = load()
    nid = max((v["id"] for v in words), default=-1) + 1
    save(words + [{"id": nid, "box": [x0, y0, x1, y1], "text": text}],
         {"op": "add", "id": nid, "box": [x0, y0, x1, y1], "new": text})
    return f"added {nid} {text}"


def main() -> None:
    ap = argparse.ArgumentParser(prog="./tool")
    sub = ap.add_subparsers(dest="cmd", required=True)
    v = sub.add_parser("view", help="image of a region + its words")
    v.add_argument("box", type=int, nargs="*", help="x0 y0 x1 y1 (default: whole page)")
    v.add_argument("--boxes", action="store_true", help="draw the word boxes")
    sub.add_parser("check", help="row arithmetic on the current words")
    f = sub.add_parser("fix", help="set a word's text; several tokens split it, '' deletes it")
    f.add_argument("id", type=int)
    f.add_argument("text")
    m = sub.add_parser("merge", help="join words that are one printed token")
    m.add_argument("ids", type=int, nargs="+")
    m.add_argument("--text")
    a = sub.add_parser("add", help="add a missed token with the pixel box of its ink")
    a.add_argument("box", type=int, nargs=4)
    a.add_argument("text")
    args = ap.parse_args(sys.argv[2:])
    if args.cmd == "view":
        if len(args.box) not in (0, 4):
            ap.error("view takes no box or x0 y0 x1 y1")
        print(view(*args.box, boxes=args.boxes))
    elif args.cmd == "check":
        print(check())
    elif args.cmd == "fix":
        print(fix_word(args.id, args.text))
    elif args.cmd == "merge":
        print(merge_words(args.ids, args.text))
    else:
        print(add_word(*args.box, args.text))


if __name__ == "__main__":
    main()
