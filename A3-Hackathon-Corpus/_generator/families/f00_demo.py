"""f00_demo: a plain DATEV-like Anlagenspiegel, assets grouped by class, class sums and a grand total.

Family contract (every families/fNN_name.py):
  * build(out: Path) -> None writes only into out/<module stem>/ (make_corpus wipes that folder first).
  * Each document: <doc_id>.xlsx or <doc_id>.pdf plus <doc_id>.key.json. doc_id = "<stem>_<nn>"; it must not
    reveal the variant, the teams see the file names.
  * Data only from core.make_client(seed, ...), a fixed seed per client: same seed, same bytes.
  * One core.Table per document: fields per column (Asset attributes, amounts in core.AMOUNTS), widths, aligns,
    header_rows of core.Cell, rows via t.add_asset / t.add_total / t.add. Every amount is int cents.
  * Render (core.render_xlsx / core.render_pdf, core.scan for a scanned copy), then core.write_key right away:
    the key takes each row's pointer from the last render of that row.
  * Problems: core.p_* change the table in place and return (table, record); pass the records to write_key.
  * Check: python3 -I _generator/make_corpus.py OUT <stem>   (builds, then runs check_corpus on OUT)
"""

import random
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import core  # noqa: E402
from core import Cell  # noqa: E402

FAMILY = Path(__file__).stem
FIELDS = ["asset_no", "name", "ahk_open", "zugang", "abgang", "umbuchung", "ahk_close",
          "kum_afa_open", "afa_gj", "afa_abgang", "zuschreibung", "kum_afa_close", "bw_close", "bw_prev"]


def table(c: core.Client) -> core.Table:
    d0, d1, prev = core.fmt_date(c.fy_start), core.fmt_date(c.fy_end), core.fmt_date(c.fy_end.replace(year=c.year - 1))
    t = core.Table(
        fields=FIELDS, col_widths=[36, 132] + [51] * 12, align=["l", "l"] + ["r"] * 12, zero="",
        title_lines=[c.company, f"Anlagenspiegel zum {d1}", f"Wirtschaftsjahr {d0} bis {d1}     Beträge in EUR"],
        header_rows=[[Cell("", 2), Cell("Anschaffungs- und Herstellungskosten", 5), Cell("Kumulierte Abschreibungen", 5),
                      Cell("Buchwerte", 2)],
                     [Cell("Anl.-Nr."), Cell("Bezeichnung"), Cell(d0), Cell("Zugänge"), Cell("Abgänge"), Cell("Umbuchungen"),
                      Cell(d1), Cell(d0), Cell("Geschäftsjahr"), Cell("Abgänge"), Cell("Zuschreib."), Cell(d1),
                      Cell(d1), Cell(prev)]],
        page_header="Seite {p} von {n}", page_footer="Druckdatum {date}", print_date=date(c.year + 1, 2, 14))
    for k, assets in c.by_class():
        t.add("heading", {0: k.account, 1: k.name})
        for a in assets:
            t.add_asset(a)
        t.add_total("class_subtotal", f"Summe {k.name}", assets, cls=k.code)
        t.add("blank", {})
    t.add_total("grand_total", "Summe Anlagevermögen", c.assets)
    return t


def build(out: Path) -> None:
    d = out / FAMILY
    d.mkdir(parents=True, exist_ok=True)
    c = core.make_client(7, year=2025, n_classes=6, assets_per_class=(3, 6))

    def key(n: int, fmt: str, variant: str, t: core.Table, planted=()) -> None:
        meta = core.key_meta(c, doc_id=f"{FAMILY}_{n:02d}", family=FAMILY, doc_type="Anlagenspiegel", format=fmt, variant=variant)
        core.write_key(d / f"{FAMILY}_{n:02d}.key.json", meta, t, planted)

    t = table(c)  # 01 perfect Excel
    core.render_xlsx(t, d / f"{FAMILY}_01.xlsx")
    key(1, "xlsx", "perfect", t)

    t = table(c)  # 02 perfect text PDF
    core.render_pdf(t, d / f"{FAMILY}_02.pdf")
    key(2, "pdf", "perfect", t)

    t = table(c)  # 03 the same print, scanned: image only
    with tempfile.TemporaryDirectory() as tmp:
        core.render_pdf(t, Path(tmp) / "print.pdf")
        core.scan(Path(tmp) / "print.pdf", d / f"{FAMILY}_03.pdf", seed=3)
    key(3, "scan", "perfect", t)

    t, rng = table(c), random.Random(4)  # 04 problems: glued numbers + a subtotal that does not tie
    planted = [core.p_glued(t, rng)[1], core.p_non_tying(t, rng)[1]]
    core.render_pdf(t, d / f"{FAMILY}_04.pdf")
    key(4, "pdf", "problem", t, planted)
