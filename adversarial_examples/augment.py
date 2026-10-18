#!/usr/bin/env python3
"""Scan-style variants of the corpus PDFs: per-page rotation (0/90/180/270), skew, paper and toner tone, noise, blur,
JPEG or fax, DPI. Every key pointer (any {page, bbox}) goes through the same geometry as its page: bbox becomes the
axis-aligned box of the turned row, quad keeps its four corners (x, y from the top left, points).

    .venv/bin/python augment.py CORPUS OUT [--n 3] [--seed 0] [--formats pdf,outline] [--only DOC_ID ...]
    .venv/bin/python augment.py --selftest

Output mirrors the corpus: OUT/<family>/<doc_id>.pdf + <doc_id>.key.json, doc_id = <family>_<nn> numbered from 51 so it
never meets the generator's own 01-50; OUT/augment.log has one line per document.

Sources are text PDFs and outline PDFs: their keys are exact. Corpus scans are skipped (already noisy, and the f06_03
key ignores its own page turns); xlsx has no page geometry. Each row box is checked to hold the ink it held before
(geometry only). Legibility after noise is not checked: no OCR engine here.
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import math
import random
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageFilter

POPPLER = "/opt/homebrew/bin"


def page_sizes(pdf: Path) -> list[tuple[float, float]]:
    info = subprocess.run([f"{POPPLER}/pdfinfo", "-f", "1", "-l", "9999", str(pdf)],
                          capture_output=True, text=True, check=True).stdout
    return [(float(w), float(h)) for w, h in re.findall(r"Page\s+\d+ size:\s+([\d.]+) x ([\d.]+)", info)]


def render(pdf: Path, dpi: int) -> list[Image.Image]:
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([f"{POPPLER}/pdftoppm", "-r", str(dpi), "-gray", "-png", str(pdf), f"{tmp}/pg"], check=True)
        return [Image.open(f).convert("L") for f in sorted(Path(tmp).glob("pg*.png"))]


def turn(im: Image.Image, angle: float):
    """Rotate counter-clockwise by `angle` degrees; the page grows so nothing is clipped. Returns the image and the
    forward map of a pixel point (the inverse of PIL's own affine)."""
    out = im.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor=255)
    a = math.radians(angle)
    c, s = math.cos(a), math.sin(a)
    (w, h), (nw, nh) = im.size, out.size

    def fwd(x: float, y: float) -> tuple[float, float]:
        dx, dy = x - w / 2, y - h / 2
        return nw / 2 + c * dx + s * dy, nh / 2 - s * dx + c * dy

    return out, fwd


def ink(im: Image.Image, box) -> int:
    x0, y0, x1, y1 = (round(v) for v in box)
    return sum(im.crop((x0, y0, x1, y1)).histogram()[:128]) if x1 > x0 and y1 > y0 else 0


def degrade(im: Image.Image, p: dict, rng: random.Random) -> Image.Image:
    lo, hi = p["toner"], p["paper"]
    im = im.point(lambda v: int(lo + v * (hi - lo) / 255))  # the white fill of a turned page becomes paper too
    if p["noise"]:
        im = Image.blend(im, Image.frombytes("L", im.size, rng.randbytes(im.width * im.height)), p["noise"])
    if p["blur"]:
        im = im.filter(ImageFilter.GaussianBlur(p["blur"]))
    if p["fax"]:  # fax 'Normal' mode: half the vertical lines, hard black/white
        w, h = im.size
        return im.resize((w, h // 2)).resize((w, h), Image.NEAREST).point(lambda v: 255 if v > p["fax"] else 0).convert("1")
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=p["jpeg"])
    return Image.open(io.BytesIO(buf.getvalue()))


def params(rng: random.Random, n: int) -> dict:
    mode = rng.choices(("upright", "sideways", "flipped_pages", "mixed"), (4, 2, 2, 2))[0]
    if mode == "upright":
        turns = [0] * n
    elif mode == "sideways":  # the whole print fed in turned, like SAP lists printed across portrait pages
        turns = [rng.choice((90, 270))] * n
    elif mode == "flipped_pages":  # single sheets upside down in the feeder
        turns = [180 if rng.random() < 0.35 else 0 for _ in range(n)]
    else:
        turns = [rng.choice((0, 0, 90, 180, 270)) for _ in range(n)]
    skew = rng.uniform(-2.5, 2.5)  # one feeder per document, a little jitter per sheet
    p = {"mode": mode, "turns": turns, "skew": [round(skew + rng.uniform(-0.4, 0.4), 2) for _ in range(n)],
         "dpi": rng.choice((120, 150, 200, 240)), "paper": rng.randint(215, 250), "toner": rng.randint(0, 70),
         "noise": round(rng.uniform(0, 0.08), 3), "blur": round(rng.uniform(0, 1.2), 2),
         "fax": rng.randint(140, 190) if rng.random() < 0.15 else 0, "jpeg": rng.randint(25, 85)}
    if p["fax"] or p["dpi"] == 120:  # ponytail: keep coarse prints readable; replace with an OCR gate once one is installed
        p["dpi"], p["blur"], p["noise"] = max(p["dpi"], 150 if not p["fax"] else 200), min(p["blur"], 0.6), min(p["noise"], 0.05)
    return p


def pointers(o, found: list) -> list:
    if isinstance(o, dict):
        if isinstance(o.get("page"), int) and isinstance(o.get("bbox"), list):
            found.append(o)
        for v in o.values():
            pointers(v, found)
    elif isinstance(o, list):
        for v in o:
            pointers(v, found)
    return found


def augment(pdf: Path, key: dict, out_dir: Path, doc_id: str, seed: str) -> list[str]:
    rng = random.Random(seed)
    sizes = page_sizes(pdf)
    p = params(rng, len(sizes))
    pages, k = render(pdf, p["dpi"]), p["dpi"] / 72
    key = copy.deepcopy(key)
    by_page: dict[int, list] = {}
    for ptr in pointers(key, []):
        by_page.setdefault(ptr["page"], []).append(ptr)
    out, problems = [], []
    for i, (im, (W, H)) in enumerate(zip(pages, sizes)):
        sx, sy = im.width / W, im.height / H
        turned, fwd = turn(im, p["turns"][i] + p["skew"][i])
        for ptr in by_page.get(i + 1, []):
            x0, t, x1, b = ptr["bbox"]
            quad = [fwd(x * sx, y * sy) for x, y in ((x0, t), (x1, t), (x1, b), (x0, b))]
            xs, ys = [x for x, _ in quad], [y for _, y in quad]
            box = (max(min(xs), 0), max(min(ys), 0), min(max(xs), turned.width), min(max(ys), turned.height))
            before, after = ink(im, (x0 * sx, t * sy, x1 * sx, b * sy)), ink(turned, box)
            if before > 20 and after < 0.9 * before:
                problems.append(f"page {i + 1} bbox {ptr['bbox']}: ink {before} -> {after}")
            ptr["bbox"] = [round(v / k, 1) for v in box]
            ptr["quad"] = [[round(x / k, 1), round(y / k, 1)] for x, y in quad]
        out.append(degrade(turned, p, rng))
    out[0].save(out_dir / f"{doc_id}.pdf", "PDF", resolution=p["dpi"], save_all=True, append_images=out[1:],
                title=None, creationDate=None, modDate=None)
    m = key["meta"]
    key["meta"] = {**m, "doc_id": doc_id, "format": "scan",
                   "augment": {"source": m["doc_id"], "source_format": m["format"], "seed": seed, **p}}
    (out_dir / f"{doc_id}.key.json").write_text(json.dumps(key, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return problems


def selftest() -> None:
    im = Image.new("L", (600, 400), 255)
    box = (100, 150, 300, 170)
    im.paste(0, box)
    for angle in (0, 90, 180, 270, 2.5, -2.5, 92.0, 268.5):
        turned, fwd = turn(im, angle)
        quad = [fwd(x, y) for x, y in ((box[0], box[1]), (box[2], box[1]), (box[2], box[3]), (box[0], box[3]))]
        aabb = (min(x for x, _ in quad), min(y for _, y in quad), max(x for x, _ in quad), max(y for _, y in quad))
        dark = sum(turned.histogram()[:128])
        assert dark > 0 and ink(turned, aabb) >= 0.98 * dark, (angle, ink(turned, aabb), dark)
        assert aabb[0] >= -1 and aabb[1] >= -1 and aabb[2] <= turned.width + 1 and aabb[3] <= turned.height + 1, angle
    print("augment selftest OK")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("corpus", nargs="?", type=Path)
    ap.add_argument("out", nargs="?", type=Path)
    ap.add_argument("--n", type=int, default=3, help="variants per source document")
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--formats", default="pdf,outline")
    ap.add_argument("--only", nargs="*", help="source doc_ids")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest or a.corpus is None:
        selftest()
        return 0
    a.out.mkdir(parents=True, exist_ok=True)
    formats, bad, i, per_family, log = set(a.formats.split(",")), 0, 0, {}, []
    for kp in sorted(a.corpus.glob("f*/*.key.json")):
        key = json.loads(kp.read_text(encoding="utf-8"))
        m = key["meta"]
        if m["format"] not in formats or (a.only and m["doc_id"] not in a.only):
            continue
        fam = m["family"]
        (a.out / fam).mkdir(exist_ok=True)
        for n in range(a.n):
            i, per_family[fam] = i + 1, per_family.get(fam, 50) + 1
            doc_id = f"{fam}_{per_family[fam]:02d}"
            problems = augment(kp.with_name(m["doc_id"] + ".pdf"), key, a.out / fam, doc_id,
                               f"aug-{a.seed}-{m['doc_id']}-{n}")
            bad += bool(problems)
            meta = json.loads((a.out / fam / f"{doc_id}.key.json").read_text(encoding="utf-8"))["meta"]["augment"]
            log.append(("FAIL " if problems else "OK   ") + f"{doc_id} <- {m['doc_id']} {meta['mode']} turns={meta['turns']} "
                       f"dpi={meta['dpi']} {'fax' if meta['fax'] else 'jpeg'}" + ("; " + "; ".join(problems[:3]) if problems else ""))
            print(log[-1])
    log.append(f"{i} documents, {i - bad} OK, {bad} FAIL")
    print(log[-1])
    (a.out / "augment.log").write_text("\n".join(log) + "\n", encoding="utf-8")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
