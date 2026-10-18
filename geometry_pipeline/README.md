# Geometry pipeline with OpenAI

End-to-end extraction for text PDFs, scanned PDFs and Excel workbooks. Code finds the geometry and extracts
values; OpenAI labels the layout and proofreads scan OCR through fixed tools. Output is one
`<doc_id>_<year>.json` per document/year, with integer-cent values and page/box or sheet/cell pointers.

From this directory:

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
# Set OPENAI_API_KEY in .env. Install the tesseract executable for scanned PDFs.
.venv/bin/python agent/run.py ../A3-Hackathon-Corpus out_openai --jobs 6
```

`--only SUBSTR` selects matching filenames. `scripts/run_openai.py` is an alias for the same entry point.
The default model is `gpt-6.1-sol`; override it with `OPENAI_MODEL`. Direct OpenAI calls use the
[Responses API](https://developers.openai.com/api/docs/guides/function-calling).
The OpenRouter route used in the original GPT experiment is also supported via `OPENAI_BASE_URL` and
`OPENAI_API_KEY`; see `.env.example`. No Claude CLI or OAuth login is needed.

1. Words with boxes (`agent/extract.py`, `agent/xlsx.py`): PDF text layer or Excel cells. Scans are deskewed,
   oriented and read with Tesseract (`agent/ocr.py`).
2. Grid (`agent/grid.py`): code groups visual lines and amount-column strips. Includes the later parser fixes
   for split/glued numbers and decimal points inside amount columns.
3. Layout (`agent/layout.py`, `agent/openai_backend.py`): one model conversation labels strips, row kinds,
   fiscal years and units. It returns JSON, never amounts. Code validates coverage and requests corrections.
4. Scan review (`agent/ocr.py`, `agent/openai_backend.py`, `agent/ocr_tools.py`): GPT sees zoomed page crops
   through `view`, checks row arithmetic with `check`, and edits OCR through `fix`, `merge` and `add`.
   Every edit is logged. A second OpenAI review inspects cells selected by document-wide column sums.
5. Records (`agent/records.py`, `agent/run.py`): code extracts values from the corrected words or Excel cells
   and writes the JSON output and layout metadata.

Scan review caches are separate from Claude results, under `ocr_cache/review_gpt/`. `OCR_JOBS` limits concurrent
OCR sessions (default 6). `MAX_CALLS` limits tool calls per review (default 300, increased from the experiment's
150 to accommodate hard fax pages). Exhausted reviews fail explicitly instead of caching incomplete results.
`OCR_AGENT=0` disables model proofreading and uses Tesseract alone; layout still uses OpenAI.

This combines the GPT tool harness from thread `b5c2e5fa-99c7-4790-bb70-aaa3a9ac89b6` with the later
parser and column-sum review fixes from `055e3d25-d5e8-4fc8-99c1-30c9551b30a8` and the shared Excel reader.
Validation after the switch: six offline regression tests pass, covering both API protocols, text PDF and Excel
extraction, actual OCR tools, cache reuse, second-pass review, retries and tool-budget exhaustion. Run them with
`.venv/bin/python -m unittest discover -s tests -v` from this directory.
Live runs through OpenRouter with `openai/gpt-6.1-sol` scored 444/444 on `f00_demo_01.xlsx`, 444/444 on
`f00_demo_02.pdf`, and 133/133 on scanned `f07_sap_class_summary_04.pdf`: 1,021 correct, no wrong, extra or missing
values. Direct OpenAI live validation is blocked by exhausted account credits; its protocol is covered offline.
These three documents do not establish whole-corpus accuracy. See `reports/openai_validation.txt`.

`out_final/` and `reports/final_text_xlsx_score.txt` are historical Claude-generated artifacts, not outputs or
measured scores of this OpenAI version. Use a fresh output directory for OpenAI. Known gaps include multiple
fiscal years on a single Excel sheet, reconciliation rows extracted as assets, and difficult scan digits.
