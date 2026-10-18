"""How reliable is raw extraction? For every printed amount in the answer keys, check that the extractor returns
exactly that string on the right page (PDF text layer or tesseract OCR) or that cell (xlsx).

Per value two checks:
  page  - the printed string is among the page's tokens (multiset: each token is used once)
  box   - additionally the token's centre lies inside the key's row box (+ margin); skipped on rotated pages

Scan pages are straightened (measured skew undone) before OCR unless --no-deskew.

Usage: .venv/bin/python scripts/check_extraction.py CORPUS_DIR OUT_JSON [--dpi 300] [--jobs 8] [--no-deskew]
       [--engine tesseract|paddle] [--format scan]
--engine paddle needs .venv-paddle (Python 3.12).
"""
import argparse
import io
import json
import re
import subprocess
import sys
from collections import Counter, defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import cv2
import numpy as np
import openpyxl
import pymupdf

from deskew_test import estimate

MARGIN = 6.0  # pt around the key's row box


AMOUNT = re.compile(r"^-?[\d.]+,\d\d-?$")


def remove_lines(gray: np.ndarray) -> np.ndarray:
    """Paint table rules white: long horizontal / vertical runs of dark pixels."""
    ink = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h, w = gray.shape
    rules = cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, w // 40), 1)))
    rules |= cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, h // 40))))
    out = gray.copy()
    out[cv2.dilate(rules, np.ones((3, 3), np.uint8)) > 0] = 255
    return out


def tesseract_words(gray: np.ndarray) -> list[tuple]:
    png = cv2.imencode(".png", gray)[1].tobytes()
    tsv = subprocess.run(["tesseract", "-", "-", "--psm", "6", "tsv"], input=png,
                         capture_output=True).stdout.decode("utf-8", "replace")
    words = []
    for line in tsv.splitlines()[1:]:
        p = line.split("\t")
        if len(p) == 12 and p[11].strip():
            x, y, w, h = map(int, p[6:10])
            words.append((x, y, x + w, y + h, p[11].strip()))
    return words


_paddle = None


def paddle_words(gray: np.ndarray) -> list[tuple]:
    """PaddleOCR (PP-OCRv5) in the same shape as tesseract_words. It boxes text regions, not words: a region
    holding several words is split at its spaces, sharing the box width by character count."""
    global _paddle
    if _paddle is None:
        from paddleocr import PaddleOCR  # only installed in .venv-paddle (no paddlepaddle wheels for Python 3.14)
        _paddle = PaddleOCR(use_doc_orientation_classify=False, use_doc_unwarping=False,
                            use_textline_orientation=False, cpu_threads=4)
    r = _paddle.predict(cv2.cvtColor(np.ascontiguousarray(gray), cv2.COLOR_GRAY2BGR))[0]
    words = []
    for text, (x0, y0, x1, y1) in zip(r["rec_texts"], r["rec_boxes"].tolist()):
        n, pos = len(text), 0
        for part in text.split():
            i = text.index(part, pos)
            pos = i + len(part)
            words.append((x0 + (x1 - x0) * i / n, y0, x0 + (x1 - x0) * pos / n, y1, part))
    return words


READERS = {"tesseract": tesseract_words, "paddle": paddle_words}


def unrotate(box: tuple, k: int, w: int, h: int) -> tuple:
    """Map a box from an image turned k*90° counter-clockwise (np.rot90) back to the original w x h image."""
    x0, y0, x1, y1 = box
    pts = [(x0, y0), (x1, y1)]
    if k == 1:
        pts = [(w - y, x) for x, y in pts]
    elif k == 2:
        pts = [(w - x, h - y) for x, y in pts]
    elif k == 3:
        pts = [(y, h - x) for x, y in pts]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def deskew(gray: np.ndarray) -> tuple[np.ndarray, float, bool]:
    """Turn the page back by its measured skew about the centre, same size: the generator skewed the clean page
    that way, so the result is in the key's coordinates. Rotations commute, so a sideways page keeps its angle;
    take the angle from whichever orientation gives the sharper line profile, and report that orientation."""
    (a0, s0), (a1, s1) = estimate(gray, with_score=True), estimate(np.rot90(gray), with_score=True)
    sideways = s1 > s0
    a = a1 if sideways else a0
    if a == 0:
        return gray, 0.0, sideways
    h, w = gray.shape
    m = cv2.getRotationMatrix2D((w / 2, h / 2), -a, 1.0)
    return cv2.warpAffine(gray, m, (w, h), flags=cv2.INTER_LINEAR, borderValue=255), a, sideways


def amounts(words: list[tuple]) -> int:
    return sum(bool(AMOUNT.match(t[4])) for t in words)


def ocr_page(png: bytes, dpi: int, lines: bool = True, straighten: bool = True,
             engine: str = "tesseract") -> tuple[list[tuple], int]:
    """Words [(x0, y0, x1, y1, text)] in PDF points of the unrotated page, and the rotation used (k*90°).
    Candidate quarter turns are read and the one with the most amount-shaped tokens is kept. A page the skew
    profile finds sideways also tries both 90° and 270° (OSD mixes those two up); otherwise the page as is plus
    OSD's proposal, which catches 180°. Tesseract always picks the turn (it is fast); another engine then reads
    the page once in that orientation."""
    gray = cv2.imdecode(np.frombuffer(png, np.uint8), cv2.IMREAD_GRAYSCALE)
    sideways = False
    if straighten:
        gray, _, sideways = deskew(gray)
        png = cv2.imencode(".png", gray)[1].tobytes()  # OSD sees the straightened page too
    if lines:
        gray = remove_lines(gray)
    h, w = gray.shape
    if sideways:
        turns = [0, 1, 3]  # the profile can call an upright page sideways (f09 p1); the amount count decides
    else:
        turns = [0]
        osd = subprocess.run(["tesseract", "-", "-", "--psm", "0"], input=png, capture_output=True)
        m = re.search(rb"Rotate: (\d+)", osd.stdout)
        if m and int(m.group(1)):
            turns.append((360 - int(m.group(1))) // 90 % 4)  # OSD: clockwise degrees; np.rot90: counter-clockwise
    k, words = max(((k, tesseract_words(np.rot90(gray, k))) for k in turns), key=lambda c: amounts(c[1]))
    if engine != "tesseract":
        words = READERS[engine](np.rot90(gray, k))
    s = 72 / dpi
    out = [tuple(c * s for c in unrotate(t[:4], k, w, h)) + (t[4],) for t in words]
    return out, k


def pdf_words(path: Path, ocr: bool, dpi: int, straighten: bool = True,
              engine: str = "tesseract") -> dict[int, tuple[list, int]]:
    doc, out = pymupdf.open(path), {}
    for i, page in enumerate(doc):
        if ocr:
            out[i + 1] = ocr_page(page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY).tobytes("png"), dpi,
                                  straighten=straighten, engine=engine)
        else:
            out[i + 1] = ([w[:5] for w in page.get_text("words")], 0)
    return out


def expected(key: dict) -> list[dict]:
    """One item per printed amount; glued pairs become one item holding the glued string."""
    items = []
    for r in key["rows"]:
        printed, flags = r.get("printed", {}), r.get("flags", {})
        glued = [f for f in printed if flags.get(f) == "glued"]
        for f, s in printed.items():
            if s and f not in glued:
                items.append({"field": f, "printed": s, "row": r, "glued": False})
        if glued:  # the right column of the pair carries both numbers
            s = "".join(printed[f] for f in glued)
            items.append({"field": "+".join(glued), "printed": s, "row": r, "glued": True})
    return items


def check_pdf(key: dict, path: Path, dpi: int, straighten: bool, engine: str) -> dict:
    fmt = key["meta"]["format"]
    pages = pdf_words(path, fmt in ("scan", "outline"), dpi, straighten and fmt == "scan", engine)
    pool = {p: Counter(w[4] for w in ws) for p, (ws, _) in pages.items()}
    digits = lambda t: re.sub(r"\D", "", t)  # noqa: E731
    dpool = {p: Counter([digits(w[4]) for w in ws] + [digits(a[4] + b[4]) for a, b in zip(ws, ws[1:])])
             for p, (ws, _) in pages.items()}
    res, misses = Counter(), []
    for it in expected(key):
        ptr = it["row"]["pointer"]
        pg, (x0, y0, x1, y1) = ptr["page"], ptr["bbox"]
        words, rot = pages[pg]
        s = it["printed"]
        res["total"] += 1
        if pool[pg][s] > 0:
            pool[pg][s] -= 1
            res["page_ok"] += 1
        else:
            near = [w[4] for w in words if y0 - MARGIN <= (w[1] + w[3]) / 2 <= y1 + MARGIN]
            misses.append({"page": pg, "field": it["field"], "want": s, "row": it["row"].get("asset_no") or
                           it["row"]["kind"], "rot": rot, "line_tokens": near[:30]})
        if dpool[pg][digits(s)] > 0:  # same digits as a token or two adjacent tokens: separators / splits only
            dpool[pg][digits(s)] -= 1
            res["digits_ok"] += 1
        if any(w[4] == s and x0 - MARGIN <= (w[0] + w[2]) / 2 <= x1 + MARGIN
               and y0 - MARGIN <= (w[1] + w[3]) / 2 <= y1 + MARGIN for w in words):
            res["box_ok"] += 1
    res["rotated_pages"] = sum(1 for _, r in pages.values() if r)
    return {"counts": dict(res), "misses": misses}


def check_xlsx(key: dict, path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    teur, res, misses = key["meta"]["unit"] == "TEUR", Counter(), []
    for r in key["rows"]:
        p = r["pointer"]
        for f, s in r.get("printed", {}).items():
            if not s:
                continue
            res["total"] += 1
            v = wb[p["sheet"]][p["cells"][f]].value
            want = abs(r["fields"][f])
            if isinstance(v, (int, float)):
                got = abs(round(v * (100_000 if teur else 100)))
            elif isinstance(v, str):
                t = v.strip().rstrip("-").lstrip("-").replace(".", "").replace(",", ".")
                got = abs(round(float(t) * (100_000 if teur else 100))) if re.fullmatch(r"\d+(\.\d+)?", t) else None
            else:
                got = None
            if got == want:
                res["page_ok"] += 1
                res["box_ok"] += 1
                res["digits_ok"] += 1
            else:
                misses.append({"cell": f"{p['sheet']}!{p['cells'][f]}", "field": f, "want": s, "got": v})
    return {"counts": dict(res), "misses": misses}


def run(args) -> tuple[str, dict]:
    keyfile, dpi, straighten, engine = args
    key = json.loads(Path(keyfile).read_text())
    doc = Path(keyfile.replace(".key.json", ".xlsx" if key["meta"]["format"] == "xlsx" else ".pdf"))
    out = check_xlsx(key, doc) if doc.suffix == ".xlsx" else check_pdf(key, doc, dpi, straighten, engine)
    out["format"], out["variant"] = key["meta"]["format"], key["meta"]["variant"]
    return key["meta"]["doc_id"], out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("corpus", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--jobs", type=int, default=8)
    ap.add_argument("--only", default="")
    ap.add_argument("--engine", choices=sorted(READERS), default="tesseract", help="OCR engine for scans and outlines")
    ap.add_argument("--format", default="", help="only documents of this format (pdf, scan, outline, xlsx)")
    ap.add_argument("--no-deskew", action="store_true", help="OCR scans without straightening them first")
    a = ap.parse_args()
    keys = [str(k) for k in sorted(a.corpus.glob("f*/*.key.json")) if a.only in k.name
            and (not a.format or json.loads(k.read_text())["meta"]["format"] == a.format)]
    with ProcessPoolExecutor(a.jobs) as ex:
        futures = [ex.submit(run, (k, a.dpi, not a.no_deskew, a.engine)) for k in keys]
        results = {}
        for f in as_completed(futures):  # report each document as it finishes, so a killed run still shows progress
            doc, r = f.result()
            results[doc] = r
            c = r["counts"]
            print(f"done {doc}: exact {c.get('page_ok', 0)}/{c['total']}", file=sys.stderr, flush=True)
    a.out.write_text(json.dumps(results, ensure_ascii=False, indent=1))
    agg = defaultdict(Counter)
    for doc, r in sorted(results.items()):
        c = r["counts"]
        agg[r["format"]].update(c)
        agg[r["format"]]["docs"] += 1
        print(f"{doc:32s} {r['format']:8s} {r['variant']:16s} exact {c.get('page_ok', 0):5d}/{c['total']:<5d} "
              f"digits {c.get('digits_ok', 0):5d}  box {c.get('box_ok', 0):5d}  rot {c.get('rotated_pages', 0)}",
              flush=True)
    print("\nformat   docs  values   exact   digits  exact+box")
    for f, c in agg.items():
        t = c["total"]
        print(f"{f:8s} {c['docs']:4d} {t:7d}  {c['page_ok'] / t:6.2%}  {c['digits_ok'] / t:6.2%}  {c['box_ok'] / t:6.2%}")


if __name__ == "__main__":
    sys.exit(main())
