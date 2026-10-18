"""Pages without a text layer (scans, outline-glyph PDFs) -> words with boxes, like the text layer gives.

scan_page (before the layout call): code straightens the page (measured skew undone, quarter turn picked by
tesseract) and tesseract supplies the words with boxes.

review_page (after the layout call, table pages only): an OpenAI agent proofreads tesseract's words against
the image through fixed function calls (ocr_tools.py: view = crops of the page, check = the row laws in code
with the value each cell would need, fix / merge / add = corrections). Code maps the boxes back to PDF points.

recheck_page (after all pages are proofread): code adds up the columns over the whole document; where a sum row is
not the sum of its rows, a second short session re-reads the cells one misread digit away from explaining it.

Proofread words are cached under ocr_cache/review_gpt/<doc_id>/p<N>.json, so reruns cost nothing.
OCR_AGENT=0: tesseract alone, no model calls.
"""
import json
import os
import sys
import threading
from pathlib import Path

import cv2
import numpy as np
import pymupdf

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from check_extraction import amounts, deskew, remove_lines, tesseract_words, unrotate  # noqa: E402
from grid import Word  # noqa: E402

DPI = 300
CACHE = ROOT / "ocr_cache" / "review_gpt"
TOOLS = Path(__file__).resolve().parent / "ocr_tools.py"
LIMIT = threading.Semaphore(int(os.environ.get("OCR_JOBS", "6")))  # OpenAI sessions at once, over all docs

PROMPT = """You proofread the OCR of one scanned page of a German fixed-asset schedule (Anlagenspiegel), already
straightened and upright ({w} x {h} px, {dpi} dpi). The page's words are tesseract's reading, with ids and pixel
boxes. Tesseract makes mistakes: wrong digits, lost or extra separators, '.' / ',' mixed up, one number split in
pieces or two numbers glued together, missed numbers, noise.

You can run only `./tool` (with Bash) and read the images it writes (with Read):
  ./tool view [x0 y0 x1 y1] [--boxes]   writes the image of a region to view_<n>.png and lists its words
                                        (id [box] text). Read the image. Use regions of about 1200 x 600 px; the
                                        full page is too small to read digits.
  ./tool check                          the row arithmetic, computed in code on the current words: rows that do
                                        not add up, their cells, and per cell the value that alone would make the
                                        row add up
  ./tool fix ID 'TEXT'                  set a word's text; 'A B' splits it in two; '' deletes it
  ./tool merge ID ID ... [--text T]     join pieces of one printed number
  ./tool add x0 y0 x1 y1 'TEXT'         add a missed token with the pixel box of its ink

1. View the whole page once, then go through the table region by region and compare every number with the
   image. Correct what tesseract misread. Copy each token exactly as printed: digits, '.', ',', '-', parentheses,
   trailing minus (1.234.567,89  1.234,56-).
2. Run check. For each row it lists, view the cells close up and decide from the ink. A hint is a candidate, not
   an answer: take it only where the image shows that value (a 3 that can be an 8, a lost digit). Where the ink
   clearly shows something else, keep what is printed; the page itself may not add up. Never change a figure
   because the sum needs it.
3. Repeat until check lists only rows whose print you have confirmed. Numbers matter most; text words (row
   names, headers) roughly right is enough.

Answer with one line: rows still listed by check and why."""


RECHECK = """You proofread this page (`page.png`, {w} x {h} px) before. Code has since added up the columns over
the whole document and found sum rows that are not the sum of their rows. These cells on this page are the ones
that could explain it with one misread digit:

{items}

Look at each of them close up and decide from the ink alone. Change a cell only where the image shows a different
figure than its current text; most of them are probably right, and a sum that does not add up can be the page's own
print. Never change a figure because the sum needs it.

You can run only `./tool` (with Bash) and read the images it writes (with Read):
  ./tool view x0 y0 x1 y1               writes the image of a region to view_<n>.png and lists its words
  ./tool fix ID 'TEXT'                  set a word's text

Answer with one line: the cells you changed."""


def straighten(page, k: int | None = None) -> tuple[np.ndarray, int, int, int]:
    """Upright, deskewed gray image, the quarter turn k used (np.rot90, ccw) and the size w, h before turning.
    k is picked by tesseract (most amounts read) unless given."""
    gray = cv2.imdecode(np.frombuffer(page.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY).tobytes("png"), np.uint8),
                        cv2.IMREAD_GRAYSCALE)
    gray, _, sideways = deskew(gray)
    h, w = gray.shape
    if k is None:
        turns = [0, 1, 3] if sideways else [0, 2]
        k = max(turns, key=lambda k: amounts(tesseract_words(remove_lines(np.rot90(gray, k)))))
    return np.ascontiguousarray(np.rot90(gray, k)), k, w, h


def to_words(px: list[dict]) -> list[Word]:
    """Pixel boxes of the upright page -> Words in PDF points of the upright page (the reading frame)."""
    s = 72 / DPI
    return [Word(*(c * s for c in x["box"]), x["text"]) for x in px if str(x["text"]).strip()]


def scan_page(page, doc_id: str) -> dict:
    """{words, to_page} like extract.pdf_pages gives for a text page, plus `scan`: what review_page and
    recheck_page need, and the cached record if the page was proofread before."""
    cache = CACHE / doc_id / f"p{page.number + 1}.json"
    record = json.loads(cache.read_text()) if cache.exists() else None
    upright, k, w, h = straighten(page, record["k"] if record else None)
    words = record["words"] if record else [{"id": i, "box": list(t[:4]), "text": t[4]}
                                            for i, t in enumerate(tesseract_words(remove_lines(upright)))]
    s = 72 / DPI

    def to_page(box):
        px = unrotate(tuple(c / s for c in box), k, w, h)
        return [round(c * s, 1) for c in px]
    return {"words": to_words(words), "to_page": to_page,
            "scan": {"png": cv2.imencode(".png", upright)[1].tobytes(), "size": upright.shape[::-1], "k": k,
                     "w": w, "h": h, "words": words, "cache": cache, "record": record}}


def ids_in(words: list[dict], box: tuple) -> list[int]:
    """Words (px) that overlap a box in PDF points of the upright page."""
    x0, y0, x1, y1 = (c * DPI / 72 for c in box)
    return [v["id"] for v in words if v["box"][0] < x1 and v["box"][2] > x0 and v["box"][1] < y1 and v["box"][3] > y0]


def _ask(scan: dict, layout: dict, prompt: str) -> tuple[str, list[dict], list, dict]:
    """One OpenAI function-calling session: answer, corrected words, edits and usage."""
    from openai_backend import ask
    return ask(scan, layout, prompt)


def _save(scan: dict) -> None:
    scan["cache"].parent.mkdir(parents=True, exist_ok=True)
    scan["cache"].write_text(json.dumps(scan["record"], ensure_ascii=False))


def review_page(scan: dict, layout: dict) -> tuple[list[Word], dict]:
    """OpenAI proofreads one scanned table page. layout: page, teur, strips, rules, kinds, headers of its group."""
    if scan["record"] or os.environ.get("OCR_AGENT", "1") == "0":
        return to_words(scan["words"]), {"cost_usd": 0.0}
    w, h = scan["size"]
    tesseract = scan["words"]
    answer, scan["words"], edits, usage = _ask(scan, layout, PROMPT.format(w=w, h=h, dpi=DPI))
    scan["record"] = {"k": scan["k"], "w": scan["w"], "h": scan["h"], "dpi": DPI, "answer": answer,
                      "usage": usage, "edits": edits, "tesseract": tesseract, "words": scan["words"]}
    _save(scan)
    return to_words(scan["words"]), usage


def recheck_page(scan: dict, layout: dict, items: list[dict]) -> tuple[list[Word], dict]:
    """Second look at the cells the document-wide column sums point to. items: {box (pt), text, row, field, sum,
    note}. Runs once per page; the result is cached with the page."""
    if not items or not scan["record"] or "recheck" in scan["record"] or os.environ.get("OCR_AGENT", "1") == "0":
        return to_words(scan["words"]), {"cost_usd": 0.0}
    lines = [f"- words {ids_in(scan['words'], it['box'])} '{it['text']}': row '{it['row'][:40]}', {it['field']}; "
             f"{it['note']}" for it in items]
    w, h = scan["size"]
    answer, scan["words"], edits, usage = _ask(scan, layout, RECHECK.format(w=w, h=h, items="\n".join(lines)))
    scan["record"].update(words=scan["words"], recheck={"items": lines, "answer": answer, "edits": edits,
                                                        "usage": usage})
    _save(scan)
    return to_words(scan["words"]), usage
