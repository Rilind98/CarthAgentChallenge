# Neurosymbolic extraction of fixed-asset schedules

Reads a client's fixed-asset schedule (Anlagenspiegel, Anlagengitter, Anlagenverzeichnis) **as a PDF**: a text PDF, a
scan, or a PDF whose glyphs are drawn as shapes. It fills the auditor's **Arbeitsversion** and writes **one JSON
record per value**: the value, the row and field it fills, its pointer (page + box), its reading confidence and a
flag when it was left blank.

The method splits the work in three:

| | who | what it may do |
|---|---|---|
| **Layout** | a vision-language model (VLM) | look at the page and say what each column means and what each row is. It answers with ids and labels only, and its answer schema has no place for a number. |
| **Figures** | code: PDF text layer or OCR | copy the printed characters inside a row × column box and parse them into cents. |
| **Verification** | symbolic: accounting laws | accept the VLM's reading only where the numbers obey the laws, find subtotals by arithmetic, and decide per value whether to write, flag or leave blank. |

So the model never writes a figure (rule 01). Every value points to its printed characters (rule 02). The laws check
both the layout and the values (rule 03). Nothing is keyed to a known layout (rule 04). A client's broken sum is
written as printed and flagged, never repaired (rule 05). Scoring is +1 correct, 0 blank, −3 wrong, so whenever
the evidence is not enough the value is left blank with a reason.

## Pipeline

```
PDF ──► 0 tokens ──► 1a geometry ──► 1b VLM labels ──► 2 extraction ──► 3 verification ──► 4 outputs
        text layer     code: columns    model: field per     code: string in       laws pick and check    Arbeitsversion.xlsx
        or OCR         from the right   column, kind per     each box → cents      the mapping; totals    05_values.json
        (+confidence)  edges of the     row, unit, period,   (+ confidence)        by arithmetic; emit /  04_doc.json
                       numbers, one     leaf level           no figures ever       flag / blank           03_detections.json
                       band per line
```

### 0 · Tokens: every printed string with its box and confidence (`pipeline/tokens.py`, `pipeline/ocr.py`)
- **Text layer** (pdfplumber), confidence 1.0. Words are read in PDF stream order, so text that overlaps a
  neighbouring column is not interleaved. Rotated pages are mapped into an upright reading frame, and pointers stay
  in page coordinates.
- **No text layer** (scans, glyph outlines): **OCR**. Each page is rendered at 300 dpi, deskewed (the angle with the
  sharpest row projection), and ruled table lines are removed by morphological filtering. Then:
  - **docTR** (DBNet detection + PARSeq recognition, on the Apple GPU) reads the page and gives a **confidence per word**;
  - pages turned by 90° are read in the orientation with the most confident text;
  - tesseract is the fallback (`OCR_ENGINE=tesseract`).
- One number style per document is decided by vote (`1.234,56` or `1,234.56`), and dates do not vote. Two amounts
  printed without a gap (`8.421,128.530,49`) are split here; the two-decimal grammar makes the split unique.

### 1a · Geometry: where the table is (code, `pipeline/layout.py`)
- **Columns** come from the right edges of the numbers. A trailing minus is ignored, and a lone `-` counts as a printed
  zero. Each column's header box runs from the previous column to its own edge. A second label in a box that names a
  field is a column printed entirely blank.
- **Rows**: one band per printed line that carries amounts. Text-only lines become headings or continue a wrapped
  name.
- **Sections**: a page that repeats the header continues the section, as does a page without a header whose
  columns line up.
- **Block layouts** (one record printed over several lines) are detected from the data and the document is left
  blank for now (see *Not built yet*).

### 1b · VLM labels: what the table means (`pipeline/vlm.py`)
The VLM (default `gpt-5.6-sol`) receives the **page image** and the geometry from 1a, with ids:
- the **columns**: header text, group header, x-position;
- the **rows**: id and label, with amounts masked as `#`.

It answers in a **strict JSON schema**:

```json
{"columns": [{"id": "c3", "field": "abgang", "confidence": 0.9}, …],
 "rows": [{"id": "r12", "kind": "subtotal"}, …],
 "unit": "EUR", "period_end": "2025-12-31", "lines_per_record": 1,
 "leaf_level": "asset", "notes": "…"}
```

- `field` is one of the 15 schedule fields or `other`.
- `kind` is one of `asset` · `subtotal` · `grand_total` · `carry_over` · `heading` · `not_schedule`.
- `leaf_level` is one of `asset` · `class` · `account` · `balance_sheet_item`.
- Every value is an id or an enum. There is no free field for a figure, and the layout is validated again afterwards
  (`schema.Layout` rejects any amount in its text).

The system prompt defines each field in German accounting terms (AHK, AfA, Buchwert × Beginn / Zugang / Abgang /
Umbuchung / Zuschreibung / Ende), how group headers and dates are read against the period, and that an undated column
means the period end.

**The VLM's answer is a proposal.** Its field becomes the leading candidate for the column, and the header lexicon's
readings stay as weaker alternatives. Step 3 decides. Rows the VLM marks `heading` or `not_schedule` are dropped.

### 2 · Extraction (code, `pipeline/extract.py`)
For each row band × column box: the token(s) inside, the printed text, signed cents, and a status:
- `value`;
- `empty` (= 0, as the brief says);
- `dash` (= 0);
- `ambiguous`, `unparsed`, `suspect_empty`.

The cell's confidence is the lowest token confidence. On OCR'd pages, an empty cell is only a zero if the page
image shows **no ink** in that box; otherwise OCR missed something and the cell becomes `suspect_empty`.

### 3 · Verification (symbolic, `pipeline/verify.py`, `pipeline/lexicon.py`)
- **Accounting laws**, in positive cents, with transfers signed:
  - AHK roll-forward;
  - accumulated-depreciation roll-forward;
  - `bw_close = ahk_close − kum_afa_close`;
  - `bw_prev = ahk_open − kum_afa_open`;
  - the net columns.

  Gaussian elimination derives exactly the laws that can be checked with the columns a document prints. For a
  net-method print, for example, it derives `bw_close = bw_prev + zugang − abgang_bw + umbuchung + zuschreibung − afa_gj`.
- **Mapping:** among all column→field hypotheses (VLM first, lexicon alternatives), the one whose laws hold on the
  rows wins. A column no law can check needs the VLM and the lexicon to agree; otherwise it is left blank
  (`field_uncertain`). A section whose laws mostly fail is left blank (`mapping_unverified`).
- **Subtotals by arithmetic:** a total is a row equal to the sum of the not-yet-summed rows before it, labelled or
  not, read top-down or bottom-up. The grand total closes everything above it and no total follows it; rows below it
  are not part of the schedule. Thousand-euro rounding is tolerated per summed row.
- **Transfers net to zero.** When they don't, and a row's cost roll-forward fails by exactly that amount, that
  transfer is blanked (`transfer_incomplete`).
- **Decisions per value:**

| situation | output |
|---|---|
| verified value | written, with pointer and confidence |
| client's subtotal differs from its rows | written **as printed** + `non_tying` |
| OCR confidence below `OCR_MIN_CONF` (default 0.75) | blank + `low_ocr_conf:<conf>` |
| OCR'd row that fails a law | that law's fields blank + `ocr_law_fail` |
| failing law points at an empty cell / a value next to it | blank + `law_implies_value` / `glued_suspect` |
| amount on the line with no column, or ink in an "empty" OCR cell | the row's empty cells blank + `suspect_empty` |

### 3b · Detection (`pipeline/detect.py`)
Names the document's irregularities in the answer keys' vocabulary, with evidence and location:
- glued numbers;
- non-tying subtotals;
- wrapped descriptions;
- carry-over lines;
- cover, blank and unrelated pages;
- a header only on the first page;
- multiple years;
- thousand-euro units;
- cut-off columns;
- broken transfers;
- unrelated tables.

### 4 · Outputs (`pipeline/output.py`, `pipeline/arbeitsversion.py`)
- `05_values.json`: one record per value:

  ```json
  {"row_id": "r1", "asset_no": "85", "field": "ahk_open", "value": 722500, "printed": "7.225,00",
   "pointer": {"page": 1, "bbox": [292.7, 208.35, 319.94, 215.35]}, "flag": null, "conf": 0.998}
  ```

- `Arbeitsversion_<id>.xlsx`: the template filled by its own field catalogs, placed per section by the leaf
  level the VLM read:
  - **asset rows** → *Anlagenverzeichnis*, Anlagen-Nr. as `100/a` and `1/1`;
  - **rows directly under a balance-sheet position** → *Gitter nach Anlagenklasse* (`Anlagenklasse`);
  - **positions** → *Anlagenklasse* (`Zwischensumme Bilanzposition`) **and** *Gitter nach Bilanzposition* (`Bilanzposition`);
  - **line-item tables** → *Gitter nach Bilanzposition*;
  - **grand total** → `Gesamtsumme` on each Gitter sheet that has rows.

  Every written value has its source as a cell comment. Columns the client does not print keep or get the
  template's formula. Every flag, blank and detection becomes a line on *Offene Punkte*.

## Run

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt   # + brew install tesseract poppler
echo "OPENAI_API_KEY=sk-..." > .env                                    # git-ignored; or export the variable
.venv/bin/python run.py <folder with PDFs> <output folder>
```

Each document gets `<output>/<id>/` with one file per step. You can follow any value through them.

| file | content |
|---|---|
| `00_tokens.json` | every string, box, confidence, source (`pdf-text` / `ocr:docTR`) |
| `01_layout_geometry.json` | columns and rows found by code |
| `01_vlm.json` | the VLM call: model, input (columns, rows), answer, tokens, seconds |
| `01_layout.json` | merged layout (no figures) |
| `02_extracted.json` | each cell: tokens, text, cents, status, confidence |
| `03_checks.json` | chosen mapping and its alternatives, laws, support per field, totals, table end |
| `03_detections.json` | irregularities found |
| `04_doc.json`, `05_values.json` | key-shaped document, one record per value |
| `Arbeitsversion_<id>.xlsx` | the filled template |
| `06_score.json` | dev only: diff against `<id>.key.json` when it lies next to the PDF |

### Configuration (environment variables)

| variable | default | |
|---|---|---|
| `OPENAI_MODEL` | `gpt-5.6-sol` | VLM for the layout labels |
| `VLM_EFFORT` | `medium` | reasoning effort of the VLM |
| `VLM` | `1` | `0` = no VLM: header lexicon + laws only |
| `OCR_ENGINE` | `doctr` | `doctr` or `tesseract` |
| `OCR_MIN_CONF` | `0.75` | values read with a lower OCR confidence are left blank |
| `OCR_DPI` | `300` | render resolution for OCR |
| `OCR_ALWAYS` | `0` | `1` = OCR even when a text layer exists |
| `OCR_LANG` / `OCR_PSM` | `eng` / `6` | tesseract only |

## Benchmarks

| tool | what it measures |
|---|---|
| `bench/score_av.py <truth dir> <run dir> --pdf-only` | the filled **Arbeitsversion** against ground-truth workbooks, cell by cell: correct / blank / incorrect (+1 / 0 / −3) |
| `bench/attribute.py <run dir>` | every blank and error assigned to its source: **layout**, **OCR**, verification, or workbook placement |
| `bench/fuzz.py <out> --seeds a-b` | **unseen layouts** with exact keys, built on the corpus generator: random header wording (German variants, abbreviations, English), column subsets and order, group headers, number styles, fonts, total placement, planted problems, text PDF / scan / outline. Develop on dev seeds, report only on held-out test seeds. |

### Results

**54 sample PDFs**, scored against the ground-truth Arbeitsversion workbooks (24,715 values), with `OCR_MIN_CONF=0.75`:

| configuration | correct | blank | incorrect | score |
|---|---|---|---|---|
| `gpt-4.1` + tesseract, first writer | 10,475 | 14,220 | 3,720 | −685 |
| `gpt-4.1` + tesseract, writer following the template conventions | 12,036 | 12,625 | 3,539 | +1,419 |
| `gpt-5.6-sol` + docTR, leaf level from the VLM | *running* | | | |

In those runs almost all "incorrect" cells are correct numbers written in rows the workbook comparison could not pair
with the ground truth (placement conventions). Only 54 cells held a different number. Against the answer keys
(key-shaped JSON), the same run had 15,032 correct values and 49 wrong.

On fuzzed unseen layouts, the VLM labelling is what generalises. The three dev documents that produced 190 wrong
values with the header lexicon alone produced 0 with VLM labels plus the law check.

## Not built yet
- **Block layouts** (one asset or class printed over several lines; the f03 and f05 families) are detected and left
  blank. The VLM already reports `lines_per_record`; the reader that maps record line *i* to its value type is missing.
- **German OCR words**: docTR's default vocabulary reads umlauts poorly in some headers. Figures are not affected.
- **Calibration**: `OCR_MIN_CONF` and the abstention rules should be tuned on the confidence/outcome pairs in
  `05_values.json` + `06_score.json`. The threshold that maximises the score is where P(correct) > 0.75.
- Excel and Word inputs are out of scope (the hand-overs are PDFs).

## Security
The API key is read from `OPENAI_API_KEY` or the git-ignored `.env`; it is never written to an output. Page images are
sent to the VLM provider, so use synthetic or approved documents only.
