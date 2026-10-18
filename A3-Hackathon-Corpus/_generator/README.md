# A3 hackathon corpus generator

Synthetic Anlagenspiegel / Anlagengitter / Anlagenverzeichnis documents (xlsx, text PDF, scan PDF, outline PDF without text layer), each with a JSON answer key. All data invented; amounts are int cents.

**Run** (system `python3`, always `-I`):
- `python3 -I _generator/make_corpus.py OUT [f00_demo ...]`: builds the families named (none named = all), then checks OUT
- `python3 -I _generator/check_corpus.py OUT`: checks only; one OK/FAIL line per document, exit 1 on any FAIL
- `python3 -I _generator/core.py` checks the data laws; `python3 -I _generator/selftest.py` runs every problem and renderer end to end

**Add a family**: copy `families/f00_demo.py` to `families/fNN_name.py`. Its docstring is the contract:
- `build(out)` writes only `out/fNN_name/<doc_id>.{xlsx,pdf}` + `<doc_id>.key.json`. The folder is wiped before each build.
- `doc_id = fNN_name_NN`. Never put the variant in the file name.
- Data: `core.make_client(seed, year=, n_classes=, ...)`, with `.years(n)` for multi-year documents.
- Layout: `core.Table(fields, col_widths, align, header_rows=[[Cell(text, colspan)]], ...)`, filled with `add_asset`, `add_total`, `add`.
- Render with `render_xlsx`, `render_pdf`, `scan` or `outline`, then call `write_key(path, key_meta(client, doc_id=, family=, doc_type=, format=, variant=), table, planted)` straight away.
- Problems: `p_*` change the table in place and return `(table, record)`. Order: `p_teur` → `p_glued` / `p_non_tying` / `p_wrap` → `p_no_repeat_header` → `p_carry_over` → `p_noise`.
- If you invent another name (supplier, person), add it to `meta["invented_names"]`. The check then searches the client folders for it.
