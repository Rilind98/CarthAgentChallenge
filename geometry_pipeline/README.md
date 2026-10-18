# Geometry pipeline with Claude Code for the model steps

Code finds the table geometry; a model only names what the geometry means and proofreads OCR. Figures always come
from code (PDF text layer, Excel cell, or OCR words corrected through fixed tools).

```
.venv/bin/python agent/run.py INPUT_DIR OUTPUT_DIR [--only SUBSTR] [--jobs 6]
.venv/bin/python scripts/score.py OUTPUT_DIR        # against truth/ made by scripts/keys_to_truth.py
```
Needs: `claude` CLI logged in (OAuth; ANTHROPIC_API_KEY is stripped), `tesseract`, Python deps in requirements.txt.

1. Words with boxes (`agent/extract.py`, `agent/xlsx.py`): PDF text layer; xlsx cells on a fixed grid (one sheet =
   one page, pointer = sheet + cell); pages without text are straightened and read by tesseract (`agent/ocr.py`).
2. Grid (`agent/grid.py`): lines x column strips, no meaning.
3. Layout (`agent/layout.py`): one headless `claude -p` call names strips, row kinds, years, unit; code validates
   the answer and asks again on problems.
4. Scanned table pages (`agent/ocr.py`, `agent/ocr_tools.py`): Claude proofreads tesseract's words; its only tool
   is `./tool view|check|fix|merge|add` - the row and column laws are computed in code, every edit is logged.
5. Records (`agent/records.py`, `agent/run.py`): one record per value, `<doc_id>_<year>.json`.

`scripts/run_openai.py` runs the same pipeline with OpenAI models in both model steps (A/B).

Results on the corpus (scorer: +1 correct, -3 wrong/extra, 0 missing): text PDFs + xlsx 24,099 / 25,676
(`reports/final_text_xlsx_score.txt`); scans see `out_final/`. Known gaps: xlsx sheets with two years on one sheet,
reconciliation blocks below xlsx registers extracted as rows.
