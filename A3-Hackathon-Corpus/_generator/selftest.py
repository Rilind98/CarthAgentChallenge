#!/usr/bin/env python3
"""End-to-end check of every problem and renderer option: builds 7 documents in a temp dir, runs check_corpus.

    python3 -I selftest.py        (data laws alone: python3 -I core.py)
"""

import random
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / "families")]
import check_corpus  # noqa: E402
import core  # noqa: E402
from f00_demo import table  # noqa: E402

with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / "st"
    out.mkdir()
    c = core.make_client(11, year=2025, fy_end=(6, 30), n_classes=9, assets_per_class=(4, 8), subnumbers=True,
                         numbering="running")
    rng = random.Random(1)

    def key(n, fmt, variant, t, planted=(), years=None):
        meta = core.key_meta(c, doc_id=f"st_{n:02d}", family="st", doc_type="Anlagenspiegel", format=fmt, variant=variant)
        core.write_key(out / f"st_{n:02d}.key.json", {**meta, **({"years": years} if years else {})}, t, planted)

    t = table(c)  # every PDF problem, in the documented order
    P = [core.p_glued(t, rng)[1], core.p_non_tying(t, rng)[1], core.p_wrap(t, rng)[1], core.p_no_repeat_header(t)[1],
         core.p_carry_over(t)[1], core.p_noise(t, c, "cover", "start")[1], core.p_noise(t, c, "blank", "middle")[1],
         core.p_noise(t, c, "unrelated", "end")[1]]
    core.render_pdf(t, out / "st_01.pdf")
    key(1, "pdf", "problem", t, P)
    assert all(P) and sum(r.kind == "carry_over" for r in t.rows) == 2

    t = table(c)  # TEUR, then a non-tying subtotal, as numbers with an offset and a noise sheet
    P = [core.p_teur(t)[1], core.p_non_tying(t, rng)[1]]
    core.render_xlsx(t, out / "st_02.xlsx", offset=(3, 2), extra_sheets={"Hinweise": [["nur Testdaten"]]})
    key(2, "xlsx", "problem", t, P)

    t = table(c)  # amounts as text, trailing minus, no merged header
    t.num_style = "de_trailing_minus"
    core.render_xlsx(t, out / "st_03.xlsx", numbers="text", merged_headers=False)
    key(3, "xlsx", "layout-mutation", t)

    ts, rec = core.p_multi_year([table(y) for y in c.years(3)])  # three years in one PDF
    core.render_pdf(ts, out / "st_04.pdf")
    key(4, "pdf", "perfect", ts, [rec], years=[2023, 2024, 2025])

    ts = [table(y) for y in c.years(2)]  # two years, one sheet each
    core.render_xlsx(ts, out / "st_05.xlsx", sheet=["2024", "2025"])
    key(5, "xlsx", "perfect", ts, [], years=[2024, 2025])

    t = table(c)  # fax scan, page 2 upside down, Arial Narrow
    t.font = "Narrow"
    core.render_pdf(t, Path(tmp) / "print.pdf")
    core.scan(Path(tmp) / "print.pdf", out / "st_06.pdf", fax=True, rotate={1: 180}, seed=6)
    key(6, "scan", "perfect", t)

    core.outline(Path(tmp) / "print.pdf", out / "st_07.pdf")  # glyphs as vector shapes, no text layer
    core.outline(Path(tmp) / "print.pdf", Path(tmp) / "again.pdf")
    assert (out / "st_07.pdf").read_bytes() == (Path(tmp) / "again.pdf").read_bytes(), "outline is not byte-stable"
    key(7, "outline", "perfect", t)
    try:  # variant is the content state only; scan and outline are formats
        key(8, "scan", "scan", t)
    except AssertionError:
        pass
    else:
        raise SystemExit("write_key took variant 'scan'")

    sys.exit(check_corpus.main(out))
