"""Synthetic scan/outline PDFs whose real Azure layout responses are recorded.

The recorded responses in ``azure_responses/`` were produced from exactly these
generated pages. Tests regenerate the pages and replay the recorded JSON through
a loopback stub; they never call the live service.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pymupdf as fitz


HERE = Path(__file__).resolve().parent


def _base_module():
    spec = importlib.util.spec_from_file_location("acceptance_fixtures_for_azure", HERE / "build_cases.py")
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _rasterise(source: Path, target: Path, *, dpi: int = 180, rotation: int = 0, cropbox: tuple[float, float, float, float] | None = None) -> None:
    # JPEG keeps each scan below the 4 MB request limit of the Azure free tier;
    # the PNG scans of build_cases.py are about 9.4 MB.
    with fitz.open(source) as doc:
        page = doc[0]
        pixels = page.get_pixmap(dpi=dpi)
        with fitz.open() as image_doc:
            new_page = image_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=pixels.tobytes("jpeg", jpg_quality=85))
            if cropbox is not None:
                new_page.set_cropbox(fitz.Rect(*cropbox))
            if rotation:
                new_page.set_rotation(rotation)
            image_doc.save(target)


def generate(directory: Path) -> Path:
    """Write the base acceptance cases plus the Azure-specific scan variants."""
    base = _base_module()
    base.generate(directory)
    _rasterise(directory / "clear.pdf", directory / "azure_scanned.pdf")
    _rasterise(directory / "clear.pdf", directory / "azure_scanned_rotated_90.pdf", rotation=90)
    _rasterise(directory / "clear.pdf", directory / "azure_scanned_cropped.pdf", cropbox=(15, 10, 832, 585))
    # Equal amounts in ten cells, as pixels only: a pointer must pick the right occurrence.
    _rasterise(directory / "repeated_values.pdf", directory / "azure_repeated_values_scanned.pdf")
    # An outline page: vector drawings only, no text layer and no images.
    with fitz.open() as doc:
        page = doc.new_page(width=842, height=595)
        page.draw_rect(fitz.Rect(30, 30, 300, 120))
        page.draw_line(fitz.Point(30, 140), fitz.Point(800, 140))
        doc.save(directory / "azure_outline_only.pdf")
    return directory

