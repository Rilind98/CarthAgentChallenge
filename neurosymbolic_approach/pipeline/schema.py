"""Pydantic models for every stage output. Each stage writes one of these as JSON.

Coordinates: `bbox` is in PDF page points, top-left origin, [x0, top, x1, bottom] (like pdfplumber and the keys).
`frame` is the same box in the upright reading frame (equal to bbox unless the page text is rotated).
"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import BaseModel, Field, model_validator

FIELDS = ("ahk_open", "zugang", "abgang", "umbuchung", "zuschreibung", "ahk_close", "kum_afa_open", "afa_gj",
          "afa_abgang", "kum_afa_close", "bw_prev", "bw_close", "abgang_bw", "bew_ahk", "bew_wb")


# ---------------------------------------------------------------- step 0: tokens (perception)

class Token(BaseModel):
    id: str
    text: str
    kind: Literal["amount", "date", "int", "dash", "text", "number", "uncached"]
    source: str  # pdf-text | xlsx | ocr:<engine>
    conf: float = 1.0
    page: int | None = None
    bbox: list[float] | None = None
    frame: list[float] | None = None
    sheet: str | None = None
    row: int | None = None
    col: int | None = None
    number: float | None = None  # xlsx numeric cell value as stored
    merged_cols: list[int] | None = None  # xlsx merged range [c0, c1]
    hidden: bool = False
    hidden_col: bool = False
    glued_from: str | None = None


class Tokens(BaseModel):
    doc_id: str
    format: Literal["pdf", "ocr-pdf", "xlsx", "image-pdf"]
    number_style: Literal["de", "en"] = "de"
    source_path: str | None = None
    pages: list[dict] = Field(default_factory=list)  # pdf: page, width, height, direction
    sheets: list[str] = Field(default_factory=list)
    tokens: list[Token]


# ---------------------------------------------------------------- step 1: layout (neural slot)

class Column(BaseModel):
    id: str
    anchor: float | None = None  # pdf: right edge of the numbers (frame x)
    span: list[float] | None = None  # pdf: frame x-range of the column
    col: int | None = None  # xlsx column index
    header: str  # header text read above the column
    group_header: str = ""
    candidates: list[tuple[str, float, str]]  # (field, score, reason)


class Section(BaseModel):
    id: str
    pages: list[int] = Field(default_factory=list)
    sheet: str | None = None
    period_end: str | None = None
    period_start: str | None = None
    unit: Literal["EUR", "TEUR"] = "EUR"
    columns: list[Column]
    id_columns: dict[str, int] = Field(default_factory=dict)  # xlsx: identity field -> column index
    header_lines: list[str] = Field(default_factory=list)
    leaf_level: str | None = None  # asset | class | account | balance_sheet_item (from the VLM)


class RowBand(BaseModel):
    id: str
    section: str
    page: int | None = None
    band: list[float] | None = None  # frame [top, bottom]
    sheet: str | None = None
    row: int | None = None
    kind_guess: Literal["asset", "total", "carry_over"]
    label: str
    context: list[str] = Field(default_factory=list)  # headings above the row, nearest last
    wrapped: bool = False  # the description continued on the next printed line


class Layout(BaseModel):
    doc_id: str
    proposer: str
    mode: Literal["line", "block", "none"]
    abstain: str | None = None
    sections: list[Section]
    rows: list[RowBand]
    notes: list[str] = Field(default_factory=list)
    page_status: dict[str, str] = Field(default_factory=dict)  # page -> schedule | continues | repeats | no_amounts | no_header | empty

    @model_validator(mode="after")
    def no_figures(self):
        """Rule 01 guard: a layout proposal may point at numbers but never contain one."""
        amount = re.compile(r"\d,\d\d(?!\d)")
        for r in self.rows:
            if amount.search(r.label) or any(amount.search(c) for c in r.context):
                raise ValueError(f"layout row {r.id} carries an amount in its label: {r.label!r}")
        for s in self.sections:
            for c in s.columns:
                if amount.search(c.header) or amount.search(c.group_header):
                    raise ValueError(f"layout column {c.id} carries an amount in its header")
        return self


# ---------------------------------------------------------------- step 2: extraction (code)

class Cell(BaseModel):
    column: str
    tokens: list[str]
    text: str
    cents: int | None  # signed, as printed
    status: Literal["value", "empty", "dash", "ambiguous", "unparsed", "uncached", "suspect_empty"]
    glued: bool = False
    conf: float = 1.0  # lowest confidence of the tokens read (text layer 1.0, OCR from the engine)
    box: list[float] | None = None  # page bbox of the printed characters
    cell: str | None = None  # xlsx coordinate


class ExtractedRow(BaseModel):
    id: str
    section: str
    kind_guess: str
    label: str
    year: int | None
    identity: dict
    cells: dict[str, Cell]
    pointer: dict


class Extracted(BaseModel):
    doc_id: str
    rows: list[ExtractedRow]


# ---------------------------------------------------------------- step 4: output (key-shaped)

class OutRow(BaseModel):
    id: str
    section: str = ""
    sums: list[str] = Field(default_factory=list)  # ids of the rows this total adds up
    kind: Literal["asset", "subtotal", "grand_total", "carry_over"]
    level: int = 0
    year: int | None
    asset_no: str | None = None
    sub_no: str | None = None
    name: str | None = None
    acq_date: str | None = None
    method: str | None = None
    label: str = ""
    context: list[str] = Field(default_factory=list)
    fields: dict[str, int | None]
    printed: dict[str, str]
    pointer: dict
    flags: dict[str, str] = Field(default_factory=dict)


class ValueRecord(BaseModel):
    doc_id: str
    row_id: str
    kind: str
    year: int | None
    asset_no: str | None
    sub_no: str | None
    label: str
    field: str
    value: int | None  # cents; None = left blank
    printed: str | None
    pointer: dict
    flag: str | None = None
    conf: float | None = None  # reading confidence of the printed characters (OCR); 1.0 for a text layer
