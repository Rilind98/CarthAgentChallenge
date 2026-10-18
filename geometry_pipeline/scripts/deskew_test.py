"""Estimate the skew of scanned pages from the image (projection profile) and compare it with the true angle,
which we can replay from the generator's seeded RNG (core.scan: rng = Random(f"scan-{seed}"); per page
optional 90/180 rotate, then rng.uniform(-skew, skew), then rng.randbytes(w*h) for the noise).

Usage: .venv/bin/python scripts/deskew_test.py
"""
import random

import cv2
import numpy as np
import pymupdf


CORPUS = "corpus/A3-Hackathon-Corpus"
# (doc, seed, skew_deg, noise, dpi, rotate) as called in the family modules
SCANS = [
    ("f00_demo/f00_demo_03", 3, 0.5, 0.02, 150, {}),
    ("f01_net_spiegel/f01_net_spiegel_01", 103, 2.5, 0.03, 150, {}),
    ("f02_asset_spiegel/f02_asset_spiegel_06", 203, 0.8, 0.02, 150, {}),  # fax
    ("f03_block_spiegel/f03_block_spiegel_03", 303, 0.5, 0.02, 150, {}),  # fax
    ("f04_verzeichnis/f04_verzeichnis_04", 404, 2.5, 0.03, 150, {}),
    ("f05_sap_class_gitter/f05_sap_class_gitter_03", 503, 0.5, 0.02, 200, {}),  # fax
    ("f06_sap_asset_gitter/f06_sap_asset_gitter_03", 603, 0.5, 0.02, 200, {2: 180, 4: 90}),  # list turned 90°
    ("f07_sap_class_summary/f07_sap_class_summary_03", 703, 2.5, 0.03, 150, {}),
    ("f08_excel_spiegel_2y/f08_excel_spiegel_2y_03", 803, 0.5, 0.02, 200, {}),
    ("f09_excel_register/f09_excel_register_03", 903, 0.5, 0.02, 200, {}),
]


def true_angles(path: str, seed: int, skew: float, noise: float, rotate: dict) -> list[float]:
    """Replay the RNG. The noise draw depends on the page size in pixels, which we read from the scan's image."""
    rng, out = random.Random(f"scan-{seed}"), []
    for i, page in enumerate(pymupdf.open(path)):
        w, h = page.get_images()[0][2:4]
        out.append(rng.uniform(-skew, skew))  # PIL: positive = counter-clockwise
        if noise:
            rng.randbytes(w * h)
    return out


def estimate(gray: np.ndarray, lo=-3.0, hi=3.0, with_score: bool = False):
    """Angle (deg, counter-clockwise as PIL) whose undo gives the sharpest horizontal ink profile."""
    small = cv2.resize(gray, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
    ink = cv2.threshold(small, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    ys, xs = np.nonzero(ink)
    h, w = ink.shape
    xs, ys = xs - w / 2, ys - h / 2

    def score(a: float) -> float:
        # project the ink pixels onto the vertical axis of the page turned back by a: no image resampling,
        # so a = 0 gets no unfair sharpness advantage
        r = np.radians(a)
        proj = xs * np.sin(r) + ys * np.cos(r)
        proj -= proj.min()
        lo_bin = np.floor(proj).astype(np.int64)
        frac = proj - lo_bin  # split each pixel between its two bins: rounding would favour a = 0
        hist = np.bincount(lo_bin, 1 - frac) + np.bincount(lo_bin + 1, frac)[: lo_bin.max() + 1]
        return float(np.var(hist) / np.mean(hist) ** 2)  # scale-free, so two orientations compare

    best, step = 0.0, 0.5
    for _ in range(4):  # coarse to fine: 0.5, 0.1, 0.02, 0.004 degrees
        grid = np.arange(max(lo, best - 5 * step), min(hi, best + 5 * step) + 1e-9, step)
        best = max(grid, key=score)
        step /= 5
    return (float(best), score(best)) if with_score else float(best)


def main() -> None:
    errs, by_kind = [], {}
    for doc, seed, skew, noise, dpi, rot in SCANS:
        path = f"{CORPUS}/{doc}.pdf"
        truth = true_angles(path, seed, skew, noise, rot)
        for i, page in enumerate(pymupdf.open(path)):
            pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
            gray = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width)
            # try both orientations; text lines give the sharper profile. Rotations commute, so a quarter turn
            # leaves the skew angle unchanged.
            (a0, s0), (a1, s1) = estimate(gray, with_score=True), estimate(np.rot90(gray), with_score=True)
            vertical = s1 > s0
            est = a1 if vertical else a0
            err = abs(est - truth[i])
            errs.append(err)
            kind = "fax" if "CCITT" in str(page.get_images()[0]) else "jpeg"
            by_kind.setdefault(kind, []).append(err)
            print(f"{doc.split('/')[1]:26s} p{i + 1}  {kind:4s} {'turned 90°' if vertical else '          '}  "
                  f"true {truth[i]:+.3f}°  estimated {est:+.3f}°  error {err:.3f}°", flush=True)
    print(f"\npages {len(errs)}  mean error {np.mean(errs):.3f}°  median {np.median(errs):.3f}°  "
          f"max {max(errs):.3f}°")
    for k, e in by_kind.items():
        print(f"  {k:4s} pages {len(e):3d}  mean {np.mean(e):.3f}°  max {max(e):.3f}°")


if __name__ == "__main__":
    main()
