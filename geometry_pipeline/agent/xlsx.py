"""Excel workbook -> pages of words with boxes, so the PDF pipeline (grid, layout, records) runs unchanged.

Each sheet is one page. Each non-empty cell is one word on a fixed grid: column c spans x = 100c .. 100c + 90,
row r spans y = 10r .. 10r + 7 (pt). Numbers are written the way a German print shows them ('1.234,56'; '1.234' for
integer formats without decimals, i.e. TEUR sheets), so grid.py finds amounts and strips as on a PDF.
Pointers go back from a box to its cell: sheet name + A1 reference.
"""
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from grid import Word

CW, RH = 100.0, 10.0  # cell pitch in pt


def cell_text(v, number_format: str) -> str:
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        neg = v < 0
        if not float(v).is_integer() or "0.0" in number_format:
            cents = round(abs(v) * 100)
            s = f"{cents // 100:,}".replace(",", ".") + f",{cents % 100:02d}"
        elif "#,#" in number_format:
            s = f"{round(abs(v)):,}".replace(",", ".")
        else:
            s = str(round(abs(v)))  # 'General' integers: ids, years, quantities
        return ("-" if neg else "") + s
    return " ".join(str(v).split())


def xlsx_pages(path: Path) -> list[dict]:
    """[{page, words, sheet, to_cell}] like extract.pdf_pages; to_cell maps a box to its A1 cell."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for i, ws in enumerate(wb.worksheets):
        words = []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                t = cell_text(c.value, c.number_format or "")
                if t:
                    words.append(Word(c.column * CW, c.row * RH, c.column * CW + 90, c.row * RH + 7, t))

        def to_cell(box):
            col = int(((box[0] + box[2]) / 2) // CW)
            row = int(((box[1] + box[3]) / 2) // RH)
            return f"{get_column_letter(max(col, 1))}{max(row, 1)}"
        out.append({"page": i + 1, "words": words, "sheet": ws.title, "to_cell": to_cell})
    return out
