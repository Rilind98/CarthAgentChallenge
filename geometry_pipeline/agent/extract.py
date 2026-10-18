"""Document -> words with boxes per page, in PDF points of the upright page.

Text PDFs: the text layer. Pages without one (scans, outline glyphs) are read by ocr.py (tesseract here; the
OpenAI proofreading of table pages comes later, in run.py). Pages whose text runs vertically (a landscape list turned onto a portrait page) are
turned so that lines run left to right; boxes stay in that reading frame and are mapped back for pointers.
"""
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf

from grid import Word


def text_direction(page) -> int:
    """Quarter turns (counter-clockwise) that make the page's text run left to right."""
    votes = {0: 0, 1: 0, 2: 0, 3: 0}
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            dx, dy = ln["dir"]
            k = 0 if dx > 0.5 else 2 if dx < -0.5 else 1 if dy < -0.5 else 3
            votes[k] += len(ln["spans"])
    return max(votes, key=votes.get)


def to_reading(x0, y0, x1, y1, k: int, w: float, h: float) -> tuple:
    """Box in page coordinates -> box in the frame where text runs left to right (page turned k*90° ccw)."""
    pts = [(x0, y0), (x1, y1)]
    if k == 1:  # text runs bottom -> top: turn the page clockwise, the left edge becomes the top
        pts = [(h - y, x) for x, y in pts]
    elif k == 2:  # upside down
        pts = [(w - x, h - y) for x, y in pts]
    elif k == 3:  # text runs top -> bottom: turn the page counter-clockwise
        pts = [(y, w - x) for x, y in pts]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def from_reading(box: tuple, k: int, w: float, h: float) -> list[float]:
    """Inverse of to_reading: reading-frame box -> page coordinates."""
    x0, y0, x1, y1 = box
    pts = [(x0, y0), (x1, y1)]
    if k == 1:
        pts = [(y, h - x) for x, y in pts]
    elif k == 2:
        pts = [(w - x, h - y) for x, y in pts]
    elif k == 3:
        pts = [(w - y, x) for x, y in pts]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return [round(min(xs), 1), round(min(ys), 1), round(max(xs), 1), round(max(ys), 1)]


def pdf_pages(path: Path) -> list[dict]:
    """[{page, k, w, h, words, to_page}] with words in the reading frame; to_page maps a box back to the page."""
    out, scans = [], []
    for i, page in enumerate(pymupdf.open(path)):
        w, h = page.rect.width, page.rect.height
        k = text_direction(page)
        words = [Word(*to_reading(*t[:4], k, w, h), t[4]) for t in page.get_text("words")]
        out.append({"page": i + 1, "k": k, "w": w, "h": h, "words": words, "pdf_page": page,
                    "to_page": lambda box, k=k, w=w, h=h: from_reading(box, k, w, h)})
        if len(words) < 5 and (page.get_images() or page.get_drawings()):  # ink but no text: scan / outline glyphs
            scans.append(out[-1])
    if scans:  # tesseract now; table pages are proofread after the layout call (ocr.review_page)
        from ocr import scan_page
        with ThreadPoolExecutor(len(scans)) as ex:
            for p, r in zip(scans, ex.map(lambda p: scan_page(p["pdf_page"], path.stem), scans)):
                p.update(r)
    return out
