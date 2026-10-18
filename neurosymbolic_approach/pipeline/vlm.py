"""Step 1 (neural) - the VLM reads the page and labels the structure that code found.

Code finds the geometry (column boxes from the numbers' right edges, one row band per printed line). The VLM gets the
page image plus those columns and rows by id and answers, in a strict JSON schema:
  - which field each column is (enum), with its own confidence,
  - what each row is (asset / subtotal / grand_total / carry_over / heading / not_schedule),
  - unit (EUR / TEUR), period end, and how many printed lines one record spans.
The schema has no place for a figure: the model can never write an amount. Its answer only becomes candidates;
step 3 checks them against the accounting laws, so a wrong label cannot turn into a wrong value unnoticed.
"""

from __future__ import annotations

import base64
import json
import os
import re
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path

from .schema import Layout

FIELDS = ["ahk_open", "zugang", "abgang", "umbuchung", "zuschreibung", "ahk_close", "kum_afa_open", "afa_gj",
          "afa_abgang", "kum_afa_close", "bw_prev", "bw_close", "abgang_bw", "bew_ahk", "bew_wb", "other"]
KINDS = ["asset", "subtotal", "grand_total", "carry_over", "heading", "not_schedule"]
MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.6-sol")
EFFORT = os.environ.get("VLM_EFFORT", "medium")

SYSTEM = """You read German fixed-asset schedules (Anlagenspiegel, Anlagengitter, Anlagenverzeichnis) for an auditor.
Code has already found the amount columns and the printed rows of a table on the page image; you label them.

Column fields (choose exactly one per column):
- ahk_open: acquisition/production cost at the start of the year (AHK/AK Beginn, Anfangsstand, Bruttowert Anfang, Vortrag, Cost opening, historische AHK)
- zugang: additions to cost in the year (Zugänge, Zugang, Additions, Investitionen)
- abgang: disposals at cost (Abgänge zu AK)
- umbuchung: transfers / reclassifications (Umbuchungen, Umgliederungen), signed
- zuschreibung: write-ups / reversals of impairment (Zuschreibungen, Wertaufholung)
- ahk_close: cost at year end (AHK Ende, Endstand, Bruttowert Ende, aktuelle AHK)
- kum_afa_open: accumulated depreciation at the start (kum. AfA Beginn, Wertberichtigung/WB Beginn)
- afa_gj: depreciation for the year (AfA des Geschäftsjahres, Normal-AfA, planmäßige AfA, Zuführung; "Zugänge" under an accumulated-depreciation header)
- afa_abgang: accumulated depreciation released on disposals (AfA auf Abgänge; "Abgänge" under an accumulated-depreciation header)
- kum_afa_close: accumulated depreciation at year end (kum. AfA Ende, Abschreibungen kumuliert)
- bw_prev: book value at the prior year end / start of year (Buchwert Vorjahr, Restbuchwert Beginn)
- bw_close: book value at year end (Buchwert/Restbuchwert 31.12., Net book value)
- abgang_bw: disposals at book value (Abgänge zum Restbuchwert; net presentation)
- bew_ahk: one net movement of cost (Bew AnschWert); bew_wb: one net movement of accumulated depreciation (Bew Wertber)
- other: anything else (capitalised interest/Zinsen, Sonder-AfA, cost centre, useful life, differences, auditor checks)
Use the group headers above the column labels and the dates against the period. Two columns never get the same field.
A column without a start marker (an opening date, Beginn, Anfang, Vortrag, Vorjahr, GJ-Beg) shows values at the period
end: an undated "AHK" in a register is ahk_close, an undated "Buchwert" is bw_close.

Row kinds: asset = one leaf line of the table (one asset, or one class/account/balance-sheet line when the table has no
asset lines); subtotal = a sum of rows (Summe, *, unlabelled sum, a group line carrying the sum of the lines below it);
grand_total = the sum of the whole table; carry_over = Übertrag/Vortrag repeating a running sum at a page break;
heading = a title or group label; not_schedule = reconciliations, notes, forecasts, anything outside the schedule.
leaf_level: asset when the leaf rows are single assets (inventory numbers, names of items); class / account when each
leaf row is an asset class or a general-ledger account; balance_sheet_item when each leaf row is a § 266 HGB line item.

Never copy or write any amount or number from the page. Answer only with the JSON schema."""


def _schema(col_ids, row_ids):
    return {"name": "layout_labels", "strict": True, "schema": {
        "type": "object", "additionalProperties": False,
        "required": ["columns", "rows", "unit", "period_end", "lines_per_record", "leaf_level", "notes"],
        "properties": {
            "columns": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                        "required": ["id", "field", "confidence"],
                        "properties": {"id": {"type": "string", "enum": col_ids}, "field": {"type": "string", "enum": FIELDS},
                                       "confidence": {"type": "number"}}}},
            "rows": {"type": "array", "items": {"type": "object", "additionalProperties": False, "required": ["id", "kind"],
                     "properties": {"id": {"type": "string", "enum": row_ids}, "kind": {"type": "string", "enum": KINDS}}}},
            "unit": {"type": "string", "enum": ["EUR", "TEUR"]},
            "period_end": {"type": "string", "description": "YYYY-MM-DD or empty"},
            "lines_per_record": {"type": "integer"},
            "leaf_level": {"type": "string", "enum": ["asset", "class", "account", "balance_sheet_item"],
                           "description": "what one leaf (non-sum) row of the table is"},
            "notes": {"type": "string", "description": "one short sentence about the layout, no numbers"}}}}


def _api_key() -> str | None:
    if os.environ.get("OPENAI_API_KEY"):
        return os.environ["OPENAI_API_KEY"]
    env = Path(__file__).resolve().parents[1] / ".env"
    if env.exists():
        for line in env.read_text().splitlines():
            if line.startswith("OPENAI_API_KEY="):
                return line.split("=", 1)[1].strip()
    return None


def available() -> bool:
    return _api_key() is not None and os.environ.get("VLM", "1") != "0"


def _page_png(pdf: Path, page: int, dpi: int = 130) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-png", str(pdf), f"{tmp}/p"], check=True)
        return base64.b64encode(next(Path(tmp).glob("p*.png")).read_bytes()).decode()


def _call(messages, schema) -> dict:
    body = {"model": MODEL, "messages": messages, "response_format": {"type": "json_schema", "json_schema": schema}}
    if MODEL.startswith(("gpt-5", "gpt-6", "o1", "o3", "o4")):  # reasoning models: default temperature only
        body["reasoning_effort"] = EFFORT
    else:
        body["temperature"] = 0
    req = urllib.request.Request("https://api.openai.com/v1/chat/completions", json.dumps(body).encode(),
                                 {"Content-Type": "application/json", "Authorization": f"Bearer {_api_key()}"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=600) as r:
                data = json.load(r)
            return {"answer": json.loads(data["choices"][0]["message"]["content"]), "usage": data.get("usage", {})}
        except Exception as e:  # transient network / rate limit
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))


def label_layout(pdf: Path, lay: Layout, pages: list[dict]) -> tuple[Layout, list[dict]]:
    """Ask the VLM to label every section; merge its answer into the layout as candidates. Returns (layout, log)."""
    log = []
    widths = {p["page"]: (p["height"] if p["direction"] != "ltr" else p["width"]) for p in pages}
    for sec in lay.sections:
        rows = [r for r in lay.rows if r.section == sec.id]
        if not rows or not sec.pages:
            continue
        page = sec.pages[0]
        W = widths.get(page, 1.0)
        cols = [{"id": c.id, "header_text": c.header, "group_header": c.group_header,
                 "x_from_pct": round(100 * c.span[0] / W, 1), "x_to_pct": round(100 * c.anchor / W, 1)} for c in sec.columns]
        row_list = [{"id": r.id, "page": r.page, "label": re.sub(r"\d,\d\d", "#", r.label)[:70]} for r in rows[:250]]
        user = (f"Section {sec.id}, pages {sec.pages}. Header lines read by code: {sec.header_lines[:6]}\n"
                f"Amount columns (left to right): {json.dumps(cols, ensure_ascii=False)}\n"
                f"Printed rows with amounts (label = the non-amount text of the line): {json.dumps(row_list, ensure_ascii=False)}\n"
                "The image shows the first page of this section. Label every column and every row.")
        messages = [{"role": "system", "content": SYSTEM},
                    {"role": "user", "content": [{"type": "text", "text": user},
                                                 {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{_page_png(pdf, page)}",
                                                                                     "detail": "high"}}]}]
        t0 = time.time()
        res = _call(messages, _schema([c.id for c in sec.columns], [r.id for r in rows[:250]]))
        ans = res["answer"]
        log.append({"section": sec.id, "model": MODEL, "seconds": round(time.time() - t0, 1), "usage": res["usage"],
                    "input": {"columns": cols, "rows": len(row_list)}, "answer": ans})
        # ---- merge: the VLM's field is the leading candidate; the lexicon's survive as weaker alternatives
        by_col = {a["id"]: a for a in ans["columns"]}
        for c in sec.columns:
            a = by_col.get(c.id)
            if not a:
                continue
            if a["field"] == "other":
                c.candidates = []
            else:
                lex_top = c.candidates[0][0] if c.candidates else None
                agree = "vlm; lexicon agrees" if lex_top == a["field"] else ("vlm; lexicon silent" if lex_top is None
                                                                              else f"vlm; lexicon says {lex_top}")
                rest = [(f, round(s * 0.4, 3), f"lexicon: {why}") for f, s, why in c.candidates if f != a["field"]]
                c.candidates = [(a["field"], round(1.0 + 0.5 * min(1.0, max(0.0, a["confidence"])), 3), agree)] + rest
        sec.leaf_level = ans.get("leaf_level")
        if ans.get("unit") == "TEUR":
            sec.unit = "TEUR"
        if not sec.period_end and re.fullmatch(r"\d{4}-\d{2}-\d{2}", ans.get("period_end", "")):
            sec.period_end = ans["period_end"]
        kinds = {a["id"]: a["kind"] for a in ans["rows"]}
        keep = []
        for r in lay.rows:
            k = kinds.get(r.id) if r.section == sec.id else None
            if k in ("heading", "not_schedule"):
                continue  # the VLM says this line is not a schedule row
            if k:
                r.kind_guess = {"asset": "asset", "subtotal": "total", "grand_total": "total", "carry_over": "carry_over"}[k]
            keep.append(r)
        lay.rows = keep
        if ans.get("lines_per_record", 1) > 1:
            lay.notes.append(f"VLM: section {sec.id} prints one record over {ans['lines_per_record']} lines")
    lay.proposer = f"geometry: code; labels: {MODEL}"
    return Layout.model_validate(lay.model_dump()), log  # re-validate: still no figure anywhere in the layout
