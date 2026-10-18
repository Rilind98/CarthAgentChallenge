"""Shared by f05/f06/f07, three SAP-like FI-AA reports of one invented ledger: client, SAP amount signs, key helper and
the boxed ALV list renderer (one- or multi-line records, cover page, page turned 90°, right edge cut off).
Not a family: make_corpus loads families/f*.py only.

Key conventions: amounts keep core's positive convention (Abgang, AfA, kumulierte AfA > 0); the print shows them as
SAP books them (credits with a trailing minus). Rows that print a value with no core field (f07 net columns) carry it
in "net_fields". Trio check:  python3 -I families/_sap_common.py OUT
"""

from __future__ import annotations

import json
import random
import sys
from collections.abc import Callable
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import core  # noqa: E402

POS7 = {"A.I.2": "2101200", "A.II.1": "2102100", "A.II.2": "2102200", "A.II.3": "2102300", "A.II.4": "2102400",
        "A.III.3": "2103300"}  # invented 7-digit SAP Bilanzposition per HGB position
NEG = frozenset({"abgang", "kum_afa_open", "afa_gj", "kum_afa_close"})  # credits: SAP prints them negative
ZERO_CLASS = core.AssetClass("5100", "0129", "Anzahlungen auf Anlagen", "A.II.4", "aib")  # a class with no asset
TRIO_SEED = 4711
TRIO = {"client_seed": TRIO_SEED,
        "docs": [f"{f}_{n:02d}" for f in ("f05_sap_class_gitter", "f06_sap_asset_gitter", "f07_sap_class_summary")
                 for n in (1, 2)],
        "ties": "same ledger and year: each f06 '*Klasse' subtotal equals the class row of f05 and of f07 "
                "(f07 nets: Bew AnschWert = zugang - abgang + umbuchung, Bew Wertber = afa_abgang + zuschreibung)"}
BLUE, LABEL, LILAC, YELLOW, GREY = (0.55, 0.71, 0.89), (0.1, 0.2, 0.55), (0.86, 0.78, 0.94), (1, 1, 0.8), (0.85,) * 3


def client(seed: int, *, fin: bool = False, **kw) -> core.Client:
    """make_client with SAP-like running asset numbers and sub-numbers. fin=False drops Beteiligungen (the only class
    with a Zuschreibung): f06 has no Zuschreibung column, so every doc that must tie to f06 drops it."""
    c = core.make_client(seed, subnumbers=True, numbering="running", **kw)
    if not fin:
        c.classes = [k for k in c.classes if k.kind != "fin"]
        c.lives = [L for L in c.lives if L.k.kind != "fin"]
        c.assets = [a for a in c.assets if a.cls in {k.code for k in c.classes}]
    return c


def ledger_classes(c: core.Client) -> list[core.AssetClass]:
    """Every class of the ledger in SAP sort order (Bilanzposition, account), plus one that never had an asset."""
    return sorted([*c.classes, ZERO_CLASS], key=lambda k: (list(core.POSITIONS).index(k.position), k.account))


def bukr(c: core.Client) -> str:
    return random.Random(f"sap-{c.seed}").choice(("3100", "4710", "5200", "6300", "7150", "8400"))


def titles(t: core.Table, c: core.Client, title: str, subtitle: str = "") -> None:
    run = c.fy_end + timedelta(days=45 + c.seed % 80)
    t.title_lines = [f"{'Berichtsdatum':>16}: {core.fmt_date(c.fy_end)}   {title}",
                     f"Erstellungsdatum: {core.fmt_date(run)}   {subtitle}".rstrip()]
    t.page_header, t.page_footer, t.print_date, t.report_date = "{p}", "", run, c.fy_end


def sums(assets: list, fields: list) -> dict:
    return {f: sum(getattr(a, f) for a in assets) for f in fields if f in core.AMOUNTS}


def amt(cents: int, t: core.Table) -> str:
    body = core.fmt(abs(cents), t.num_style, t.zero)
    if cents >= 0 or body == "0":
        return body
    return {"trailing": body + "-", "leading": "-" + body, "paren": f"({body})"}[getattr(t, "neg", "trailing")]


def finish(t: core.Table) -> core.Table:
    """Print every amount the SAP way (credits negative, sign style t.neg), 'nil' columns as zero, 'curr' as the unit.
    Call after building and again after each core.p_* problem: they print amounts unsigned."""
    for r in t.rows:
        if not r.truth or "fields" not in r.truth:
            continue
        tf, net = r.truth["fields"], r.truth.get("net_fields", {})
        for i, f in enumerate(t.fields):
            if r.flags.get(f) == "glued":
                continue
            if f in tf:
                r.cells[i] = r.truth["printed"][f] = amt(-tf[f] if f in NEG else tf[f], t)
            elif f in net:
                net[f] = core.round_teur(net[f]) if t.unit == "TEUR" else net[f]
                r.cells[i] = amt(net[f], t)
            elif f == "nil":
                r.cells[i] = amt(0, t)
            elif f == "curr":
                r.cells[i] = t.unit
    return t


def write_key(d: Path, family: str, n: int, c: core.Client, fmt: str, variant: str, tables, planted=(), *,
              trio: bool = False, **extra) -> None:
    doc_id = f"{family}_{n:02d}"
    meta = core.key_meta(c, doc_id=doc_id, family=family, doc_type="Anlagengitter", format=fmt, variant=variant,
                         software="SAP-like FI-AA list print", company_code=bukr(c),
                         position_codes={v: k for k, v in POS7.items()}, **extra)
    if trio:
        meta["trio"] = TRIO
    core.write_key(d / f"{doc_id}.key.json", meta, tables, planted)


# ============================================================================ renderer


@dataclass
class Look:
    font: str = "Courier"
    size: float = 7.0
    W: float = 842.0  # landscape content box
    H: float = 595.0
    rotate: bool = False  # content turned 90° counter-clockwise onto a portrait page (reads bottom to top)
    x0: float | None = None  # left edge of the table; None = centred
    widths: list | None = None  # slot widths; None = t.col_widths
    lines: list | None = None  # record lines: field index (or None) per slot; None = one line, slot i = field i
    heads: list | None = None  # heading text per line and slot; None = t.header_rows[0]
    head_fill: tuple | None = None
    label_color: tuple = (0, 0, 0)
    box: str = "grid"  # grid: every cell · record: each row · group: runs of rows with one group(row)
    group: Callable | None = None
    style: Callable | None = None  # row -> (block fill, [(line, slot_from, slot_to, fill)], bold)
    block: Callable | None = None  # row -> (labels, values): group block under the page header
    cover: Callable | None = None  # table -> boxed tables (rows of str, first = heading) for the selection page
    cover_fill: tuple | None = None
    cut: int = 0  # characters lost at the right edge of the table (printed past the page edge)
    gap: float = 0.0  # space between boxes


def render(tables, path, look: Look) -> None:
    """SAP list print. Pages follow core._plan, so core.p_* problems land where they land in core.render_pdf.
    Sets row.pointer = {page, line, bbox} like core.render_pdf (bbox in page points, top-left origin, after the turn)."""
    c = canvas.Canvas(str(path), invariant=1, pageCompression=1)
    page = 0
    for t in tables if isinstance(tables, list) else [tables]:
        page = _pages(c, t, look, page)
    c.save()


def _pages(c, t: core.Table, L: Look, page: int) -> int:
    size, lh, bold = L.size, L.size * 1.3, L.font + "-Bold"
    cw = stringWidth("0", L.font, size)
    widths = L.widths or t.col_widths
    lines = L.lines or [list(range(len(t.fields)))]
    heads = L.heads or [[h.text for h in t.header_rows[0]]]
    x0 = (L.W - sum(widths)) / 2 if L.x0 is None else L.x0
    xs = [x0 + sum(widths[:i]) for i in range(len(widths) + 1)]
    x1, cut = xs[-1], (xs[-1] - L.cut * cw if L.cut else None)
    pad = getattr(t, "neg", "trailing") != "leading"  # positives keep the sign place free: digits line up

    def start():
        c.setPageSize((L.H, L.W) if L.rotate else (L.W, L.H))
        if L.rotate:
            c.translate(L.H, 0)
            c.rotate(90)
        c.setLineWidth(0.5)

    def bbox(a, top, b, bot):  # local box, bottom-left origin -> page [x0, top, x1, bottom], top-left origin
        if L.rotate:
            return [round(L.H - top, 1), round(L.W - b, 1), round(L.H - bot, 1), round(L.W - a, 1)]
        return [round(a, 1), round(L.H - top, 1), round(b, 1), round(L.H - bot, 1)]

    def put(s, x, w, align, y, font):
        sw = stringWidth(s, font, size)
        xa = x + w - 2 - sw if align == "r" else x + (w - sw) / 2 if align == "c" else x + 2
        while cut is not None and s and xa + stringWidth(s, font, size) > cut:
            s = s[:-1]  # past the page edge: lost, also for text extraction
        if s:
            c.setFont(font, size)
            c.drawString(xa, y, s)

    def fill(a, bot, b, top, rgb):
        c.setFillColorRGB(*rgb)
        c.rect(a, bot, b - a, top - bot, stroke=0, fill=1)
        c.setFillColorRGB(0, 0, 0)

    def frame(top, bot, grid):
        c.rect(x0, bot, x1 - x0, top - bot)
        for x in xs[1:-1] if grid else ():
            c.line(x, bot, x, top)

    def cover(tabs):
        y = L.H - 34
        for tab in tabs:
            n = max(map(len, tab))
            ws = [max(len(r[j]) for r in tab if j < len(r)) * cw + 8 for j in range(n)]
            xx = [20 + sum(ws[:j]) for j in range(n + 1)]
            lo, hi = y - 0.3 * lh - 1, y + 0.7 * lh + 1
            for j, s in enumerate(tab[0]):
                if L.cover_fill:
                    fill(xx[j] + 1.5, lo + 1.5, xx[j + 1] - 1.5, hi - 1.5, L.cover_fill)
                c.rect(xx[j], lo, ws[j], hi - lo)
                put(s, xx[j], ws[j], "l", y, bold)
            y -= lh + 6
            hi = y + 0.7 * lh + 2
            for row in tab[1:]:
                for j, s in enumerate(row):
                    put(s, xx[j], ws[j], "r" if tab[0][j] == "Anzahl" else "c" if s == "X" and tab[0][j] != "Wert" else "l", y, L.font)
                y -= lh
            lo = y + 0.7 * lh - 2
            c.rect(xx[0], lo, xx[-1] - xx[0], hi - lo)
            for x in xx[1:-1]:
                c.line(x, lo, x, hi)
            y -= lh + 6

    tp, first = 0, True
    for p in core._plan(t):
        if isinstance(p, core.Row):  # noise page, drawn as core.render_pdf draws it, on a portrait sheet
            page += 1
            c.setPageSize(A4)
            c.setFont("Helvetica", 10)
            for i, s in enumerate(p.cells):
                left, _, right = s.partition("\t")
                c.drawString(70, A4[1] - 90 - i * 14, left)
                c.drawRightString(480, A4[1] - 90 - i * 14, right)
            p.pointer = {"page": page, "line": 1, "bbox": [0, 0, round(A4[0], 1), round(A4[1], 1)]}
            c.showPage()
            continue
        if first and L.cover:  # selection page = page 1 of the run, unnumbered
            page, tp = page + 1, 1
            start()
            cover(L.cover(t))
            c.showPage()
        page, tp, show, first = page + 1, tp + 1, first or t.repeat_header, False
        start()
        y = L.H - 24
        c.setFont(L.font, size)
        c.drawRightString(cut or x1, y - lh, str(tp))
        if cut is not None:
            c.saveState()
            clip = c.beginPath()
            clip.rect(0, 0, cut, L.H)
            c.clipPath(clip, stroke=0, fill=0)
        if show:
            for s in t.title_lines:
                lab, sep, rest = s.partition(": ")
                c.setFillColorRGB(*L.label_color)
                c.setFont(bold, size)
                c.drawString(x0, y, lab + sep)
                c.setFillColorRGB(0, 0, 0)
                c.setFont(L.font, size)
                c.drawString(x0 + stringWidth(lab + sep, bold, size), y, rest)
                y -= lh
            c.line(x0, y + 0.4 * lh, x1, y + 0.4 * lh)
            y -= 0.9 * lh
            r0 = next((r for r in p if r.truth), None)
            if L.block and r0:
                labels, values = L.block(r0)
                xx = [x0]
                for a, b in zip(labels, values):
                    xx.append(xx[-1] + (max(len(a), len(b)) + 2) * cw)
                for j, (a, b) in enumerate(zip(labels, values)):
                    put(a, xx[j], 0, "l", y, bold)
                    put(b, xx[j], 0, "l", y - lh, L.font)
                c.rect(x0, y - 1.3 * lh - 1, xx[-1] - x0, 2 * lh + 2)
                y -= 2.8 * lh
            top = y + 0.7 * lh + 1
            for li, head in enumerate(heads):
                if L.head_fill:
                    fill(x0, y - 0.3 * lh, x1, y + 0.7 * lh, L.head_fill)
                for si, s in enumerate(head):
                    fi = lines[li][si] if li < len(lines) else None
                    if s:
                        put(s, xs[si], widths[si], "l" if fi is None else t.align[fi], y, bold)
                y -= lh
            frame(top, y + 0.7 * lh - 1, L.box == "grid")
            y -= L.gap
        y0, gkey, gtop, gbot = y, None, None, 0.0
        for r in p:
            if L.box == "group":
                k = L.group(r)
                if gtop is not None and k != gkey:
                    frame(gtop, gbot, False)
                    gtop, y = None, y - L.gap
                gkey = k
            top, bot = y + 0.7 * lh, y - (len(lines) - 1) * lh - 0.3 * lh
            if bot < 12:
                raise ValueError(f"rows_per_page={t.rows_per_page} overflows the page")
            blk, cells, strong = L.style(r) if L.style else (None, (), False)
            if blk:
                fill(x0, bot, x1, top, blk)
            for li, sa, sb, rgb in cells:
                fill(xs[sa], y - li * lh - 0.3 * lh, xs[sb + 1], y - li * lh + 0.7 * lh, rgb)
            for li, line in enumerate(lines):
                for si, fi in enumerate(line):
                    s = r.cells[fi] if fi is not None else ""
                    if s:
                        s += " " if pad and t.align[fi] == "r" and s[-1].isdigit() else ""
                        put(s, xs[si], widths[si], t.align[fi], y - li * lh, bold if strong else L.font)
            if L.box == "group":
                gtop, gbot = top if gtop is None else gtop, bot
            else:
                frame(top, bot, L.box == "grid")
            r.pointer = {"page": page, "line": round((y0 - y) / lh) + 1, "bbox": bbox(x0, top, x1, bot)}
            y -= len(lines) * lh + (L.gap if L.box == "record" else 0)
        if gtop is not None:
            frame(gtop, gbot, False)
        if cut is not None:
            c.restoreState()
        c.showPage()
    return page


# ============================================================================ trio check


def trio_check(out: Path) -> None:
    """Class totals of the trio tie: f06 '*Klasse' subtotals = f05 class rows = f07 class rows (f07 nets derived)."""
    keys = {d.split("_")[0]: json.loads((out / d.rsplit("_", 1)[0] / f"{d}.key.json").read_text("utf-8"))
            for d in TRIO["docs"][::2]}
    f06 = {r["class"]: r["fields"] for r in keys["f06"]["rows"] if r["kind"] == "class_subtotal"}
    for fam in ("f05", "f07"):
        rows = {r["class"]: r for r in keys[fam]["rows"] if r.get("level") == "class"}
        assert set(f06) <= set(rows), (fam, set(f06) - set(rows))
        for cls, r in rows.items():
            w = f06.get(cls, {})  # the zero class has no f06 rows: every value must be 0
            w = {**w, "bew_ahk": w.get("zugang", 0) - w.get("abgang", 0) + w.get("umbuchung", 0),
                 "bew_wb": w.get("afa_abgang", 0)}
            bad = {f: (v, w.get(f, 0)) for f, v in {**r["fields"], **r.get("net_fields", {})}.items() if v != w.get(f, 0)}
            assert not bad, (fam, cls, bad)
    print(f"trio OK: {len(f06)} classes tie across f05, f06, f07 (+ {ZERO_CLASS.code} all zero)")


if __name__ == "__main__":
    trio_check(Path(sys.argv[1]))
