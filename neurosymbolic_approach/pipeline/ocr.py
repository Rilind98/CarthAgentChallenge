"""OCR for PDFs without a usable text layer (scans, glyph outlines): tesseract, word boxes with confidence.

Each page is rendered at OCR_DPI, read by tesseract (TSV output: one line per word with its confidence 0-100), and the
boxes are converted to PDF points so pointers are in the same coordinates as the text layer. A page turned by 90° is
read in the orientation that gives the best confidence; its boxes are mapped back to the page.
Confidence is kept on every token (0..1) so the decision threshold can be tuned later (config OCR_MIN_CONF).
"""

from __future__ import annotations

import csv
import io
import os
import re
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

OCR_DPI = int(os.environ.get("OCR_DPI", "300"))
OCR_LANG = os.environ.get("OCR_LANG", "eng")
PSM = os.environ.get("OCR_PSM", "6")


def _deskew(a):
    """Small skew of a scan: the angle at which the text rows give the sharpest horizontal projection."""
    import cv2
    import numpy as np
    small = cv2.resize(a, (a.shape[1] // 4, a.shape[0] // 4))
    ink = (small < 128).astype(np.float32)
    best, best_angle = -1.0, 0.0
    for ang in np.arange(-2.0, 2.01, 0.1):
        M = cv2.getRotationMatrix2D((ink.shape[1] / 2, ink.shape[0] / 2), ang, 1.0)
        prof = cv2.warpAffine(ink, M, (ink.shape[1], ink.shape[0])).sum(axis=1)
        score = float(np.var(prof))
        if score > best:
            best, best_angle = score, float(ang)
    if abs(best_angle) < 0.05:
        return a, 0.0
    M = cv2.getRotationMatrix2D((a.shape[1] / 2, a.shape[0] / 2), best_angle, 1.0)
    return cv2.warpAffine(a, M, (a.shape[1], a.shape[0]), borderValue=255), best_angle


def _clean(img: Image.Image) -> Image.Image:
    """Straighten the scan, remove ruled table lines (they read as letters), binarise for OCR."""
    import cv2
    import numpy as np
    a, _ = _deskew(np.array(img.convert("L")))
    bw = cv2.adaptiveThreshold(a, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15)
    lines = np.zeros_like(bw)
    for k in ((max(40, a.shape[1] // 30), 1), (1, max(40, a.shape[0] // 30))):  # long horizontal / vertical strokes
        lines |= cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, k))
    lines = cv2.dilate(lines, np.ones((3, 3), np.uint8))
    out = 255 - bw
    out[lines > 0] = 255
    return Image.fromarray(out)


OCR_ENGINE = os.environ.get("OCR_ENGINE", "doctr")  # doctr (deep model, default) | tesseract
_DOCTR = None


def _doctr(img: Image.Image) -> list[dict]:
    """docTR: DBNet text detection + PARSeq recognition; one confidence per word."""
    global _DOCTR
    import numpy as np
    if _DOCTR is None:
        import torch
        from doctr.models import ocr_predictor
        _DOCTR = ocr_predictor(det_arch="db_resnet50", reco_arch="parseq", pretrained=True,
                               assume_straight_pages=True, preserve_aspect_ratio=True)
        if torch.backends.mps.is_available():
            _DOCTR = _DOCTR.to("mps")
    a = np.array(img.convert("RGB"))
    page = _DOCTR([a]).pages[0]
    h, w = a.shape[:2]
    words = []
    for block in page.blocks:
        for line in block.lines:
            for wd in line.words:
                (x0, y0), (x1, y1) = wd.geometry
                text = wd.value.strip().strip("|![]{}“”\"'`‘’_")
                if text and re.search(r"\w", text):
                    words.append({"text": text, "x0": x0 * w, "top": y0 * h, "x1": x1 * w, "bottom": y1 * h,
                                  "conf": round(float(wd.confidence), 3)})
    return words


def _read(img: Image.Image) -> list[dict]:
    if OCR_ENGINE == "doctr":
        try:
            return _doctr(_clean(img).convert("RGB"))
        except ImportError:
            pass
    return _tesseract(img)


def _tesseract(img: Image.Image) -> list[dict]:
    img = _clean(img)
    with tempfile.NamedTemporaryFile(suffix=".png") as f:
        img.save(f.name)
        out = subprocess.run(["tesseract", f.name, "-", "--psm", PSM, "-l", OCR_LANG, "tsv"],
                             capture_output=True, text=True, check=True).stdout
    words = []
    for r in csv.DictReader(io.StringIO(out), delimiter="\t", quoting=csv.QUOTE_NONE):
        if r.get("level") == "5" and r.get("text", "").strip() and float(r.get("conf", -1)) >= 0:
            x, y, w, h = (int(r[k]) for k in ("left", "top", "width", "height"))
            text = r["text"].strip().strip("|![]{}“”\"'`‘’_")  # remains of ruled lines stick to word ends
            if not text or not re.search(r"[\w]", text):
                continue
            words.append({"text": text, "x0": x, "top": y, "x1": x + w, "bottom": y + h,
                          "conf": round(float(r["conf"]) / 100, 3)})
    return words


def _score(words: list[dict]) -> float:
    return sum(w["conf"] for w in words)  # many confident words = the right orientation


def ocr_pdf(path: Path) -> tuple[list[dict], dict[int, list[dict]]]:
    """Returns (pages, page_words): pages carry width/height in points and the reading direction; words are in page
    points (top-left origin), with conf in 0..1."""
    pages, page_words = [], {}
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(OCR_DPI), "-gray", "-png", str(path), f"{tmp}/pg"], check=True)
        for pno, f in enumerate(sorted(Path(tmp).glob("pg*.png")), 1):
            im = Image.open(f)
            s = 72 / OCR_DPI
            W, H = im.width * s, im.height * s
            best = ("ltr", _read(im), im.width, im.height)
            if _score(best[1]) < 40 or sum(w["conf"] >= 0.6 for w in best[1]) < 0.5 * max(1, len(best[1])):
                for direction, angle in (("btt", -90), ("ttb", 90)):  # text running up / down the page
                    r = im.rotate(angle, expand=True)
                    words = _read(r)
                    if _score(words) > _score(best[1]) * 1.3:
                        best = (direction, words, r.width, r.height)
            direction, words, _, _ = best
            out = []
            for w in words:
                u0, v0, u1, v1 = w["x0"] * s, w["top"] * s, w["x1"] * s, w["bottom"] * s  # upright image, points
                if direction == "btt":  # upright (u, v) = (H - y, x)
                    box = (v0, H - u1, v1, H - u0)
                elif direction == "ttb":  # upright (u, v) = (y, W - x)
                    box = (W - v1, u0, W - v0, u1)
                else:
                    box = (u0, v0, u1, v1)
                out.append({"text": w["text"], "x0": box[0], "top": box[1], "x1": box[2], "bottom": box[3], "conf": w["conf"]})
            pages.append({"page": pno, "width": round(W, 2), "height": round(H, 2), "direction": direction,
                          "chars": 0, "ocr_words": len(out)})
            page_words[pno] = out
    return pages, page_words


_INK_CACHE: dict = {}


def ink_ratio(pdf: str, page: int, box: list[float], dpi: int = 100) -> float:
    """Share of dark pixels inside a box (page points), ruled lines removed: is anything printed there?"""
    import numpy as np
    key = (pdf, page)
    if key not in _INK_CACHE:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-gray", "-png", pdf, f"{tmp}/p"], check=True)
            img = Image.open(next(Path(tmp).glob("p*.png")))
            _INK_CACHE[key] = np.array(_clean(img))
    a = _INK_CACHE[key]
    s = dpi / 72
    x0, y0, x1, y1 = (int(round(v * s)) for v in box)
    x0, y0, x1, y1 = max(0, x0 + 1), max(0, y0 + 1), min(a.shape[1], x1 - 1), min(a.shape[0], y1 - 1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    return float((a[y0:y1, x0:x1] < 128).mean())
