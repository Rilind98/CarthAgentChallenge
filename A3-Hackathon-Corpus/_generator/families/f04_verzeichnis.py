"""f04_verzeichnis: tax-view asset register "Anlageverzeichnis steuerrechtlich" (look-alike of profiles qsxqsx and
qsxqsxq). Portrait, Konto > Anlage, one asset per row: AHK and Buchwert at period end, acquisition date, AfA code,
ND in years. "Summe" per Konto, no grand total, no carry-over; groups run across page breaks without repeating the
Konto line. Disposed assets stay listed with AHK and Buchwert 0,00 and a DD.MM.YY disposal date. Long descriptions
wrap inside their cell. Printed with every cell ruled and a company block top right, or with no lines at all and a
redaction box top right that clips the end of the period line (image and text). Contract: see f00_demo.
"""

import random
import sys
import tempfile
from dataclasses import replace
from datetime import timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import core  # noqa: E402

FAMILY = Path(__file__).stem
# (column key, heading lines, width pt, align)
COLS = [("asset_no", ("Inventar-", "Nummer"), 65, "l"), ("name", ("Inventar-", "Bezeichnung"), 127, "l"),
        ("kst", ("Kosten-", "stelle"), 31, "r"), ("ahk_close", ("AHK", "EUR"), 78, "r"),
        ("acq_date", ("Anschaff.", "Datum"), 42, "l"), ("afa", ("AfA", "Art"), 20, "l"), ("nd", ("ND", ""), 20, "l"),
        ("bw_close", ("Buchwert", "EUR"), 90, "r"), ("disp", ("Datum des", "Abgangs"), 36, "l")]
MUT = [("asset_no", ("Anlage-", "Nr."), 50, "l"), ("name", ("Bezeichnung", ""), 150, "l"),
       ("acq_date", ("Zugang", "am"), 44, "l"), ("nd", ("Nutzungs-", "dauer"), 34, "r"),
       ("afa", ("AfA-", "Methode"), 34, "l"), ("ahk_close", ("Anschaffungs-", "wert EUR"), 70, "r"),
       ("bw_close", ("Rest-", "buchwert EUR"), 70, "r"), ("disp", ("Abgang", "am"), 40, "l"), ("kst", ("KSt", ""), 24, "r")]
KEYED = ("asset_no", "name", "ahk_close", "acq_date", "bw_close")  # Table fields; the other columns are set by hand
X0, BOTTOM, LINE = 43, 805, 8.6


def table(c: core.Client, cols: list, kst: dict, period: str, *, num_style="de", konto="Konto  {acc}") -> core.Table:
    keys = [k for k, *_ in cols]
    t = core.Table(fields=[k if k in KEYED else None for k in keys], col_widths=[w for *_, w, _ in cols],
                   align=[a for *_, a in cols], header_rows=[[core.Cell(h[i]) for _, h, *_ in cols] for i in (0, 1)],
                   title_lines=["Anlageverzeichnis steuerrechtlich", period, "Anlagegruppe: A. Anlagevermögen"],
                   num_style=num_style, font_size=7.5, landscape=False, page_footer="Seite: {p}",
                   print_date=c.fy_end + timedelta(days=60 + c.seed % 120))
    for k, assets in c.by_class():
        acc = assets[0].account
        t.add("blank", {})
        t.add("heading", {"asset_no": konto.format(acc=acc, name=k.name)})
        for a in sorted(assets, key=lambda a: a.asset_no):  # text sort, as the source prints them
            r = t.add_asset(a)
            extra = {"kst": kst[k.code], "nd": str(a.nd_months // 12), "disp": core.fmt_date(a.disposal_date, "de_short"),
                     "afa": {"GWG": "sb", "keine": "k"}.get(a.method) or ("geb" if a.position == "A.II.1" else "lin")}
            for i, key in enumerate(keys):
                r.cells[i] = extra.get(key, r.cells[i])
        r = t.add_total("account_subtotal", "", assets, cls=k.code, account=acc)
        r.cells[keys.index("asset_no")] = "Summe"
    return t


def wrap(text: str, width: float, size: float) -> list:
    lines, cur = [], ""
    for w in text.split(" "):
        if cur and stringWidth(f"{cur} {w}", "Helvetica", size) > width:
            lines, cur = lines + [cur], w
        else:
            cur = f"{cur} {w}" if cur else w
    return lines + [cur]


def render_pdf(t: core.Table, path: Path, *, ruled: bool, company: list | None) -> None:
    """Text PDF, portrait A4. ruled: every cell boxed. company: 4 lines top right; None = redaction box that clips the
    end of the period line. Sets each row's pointer {page, line, bbox} (top-left origin points)."""
    W, H = A4
    fs, ni = t.font_size, t.fields.index("name")
    xs = [X0 + sum(t.col_widths[:i]) for i in range(len(t.col_widths) + 1)]
    reg, bold = "Helvetica", "Helvetica-Bold"
    cv = canvas.Canvas(str(path), invariant=1, pageCompression=1)
    st = {"page": 0, "tp": 0, "y": None, "n": 0}

    def text(s, x, y, font=reg, size=fs, align="l", **kw):
        cv.setFont(font, size)
        if align == "l":
            cv.drawString(x, H - y, s, **kw)
        else:
            (cv.drawRightString if align == "r" else cv.drawCentredString)(x, H - y, s)

    def cell(s, i, y, font=reg):
        a = t.align[i]
        text(s, {"l": xs[i] + 2, "r": xs[i + 1] - 2, "c": (xs[i] + xs[i + 1]) / 2}[a], y, font, align=a)

    def grid(top, bot, a=0, b=len(t.col_widths)):
        if ruled:
            cv.setLineWidth(0.5)
            cv.rect(xs[a], H - bot, xs[b] - xs[a], bot - top)
            for x in xs[a + 1:b]:
                cv.line(x, H - top, x, H - bot)

    def close():
        if st["y"] is not None:
            cv.showPage()
            st["y"] = None

    def open_page(head):
        close()
        st["page"] += 1
        st["tp"] += 1
        st["n"] = 0
        cv.setPageSize((W, H))
        if ruled:
            cv.setLineWidth(0.4)
            cv.line(xs[0], H - 814, xs[-1], H - 814)
        text(f"Datum: {core.fmt_date(t.print_date, 'de_short')}", xs[0], 826, size=7)
        text(t.page_footer.format(p=st["tp"]), xs[-1], 826, size=7, align="r")
        st["y"] = 40
        if head:
            if ruled:
                text(t.title_lines[0], xs[0], 46, size=10, charSpace=0.8)
            else:
                text(t.title_lines[0], xs[0], 46, bold, 10)
            for i, s in enumerate(company or []):
                text(s, xs[-1], 40 + i * 9, align="r")
            period = t.title_lines[1]
            x = (W - stringWidth(period, reg, 8)) / 2
            if company:
                text(period, x, 86, size=8)
            else:  # the redaction box clipped the last two digits of the year, in the image and the text layer
                text(period[:-2], x, 86, bold, 8)
                bx = x + stringWidth(period[:-2], bold, 8) + 0.5
                cv.setFillColorRGB(0, 0, 0)
                cv.rect(bx, H - 90, xs[-1] + 4 - bx, 68, stroke=0, fill=1)
                cv.setStrokeColorRGB(0.65, 0.65, 0.65)
                cv.setLineWidth(0.7)
                cv.line(bx, H - 90, xs[-1] + 4, H - 22)
                cv.line(bx, H - 22, xs[-1] + 4, H - 90)
                cv.setStrokeColorRGB(0, 0, 0)
            text(t.title_lines[2], W / 2, 95, bold, 8, align="c")
            for j, hr in enumerate(t.header_rows):
                for i, h in enumerate(hr):
                    text(h.text, (xs[i] + xs[i + 1]) / 2, 113 + 9 * j, reg if ruled else bold, 7, align="c")
            grid(103, 128)
            st["y"] = 128

    first, gap = True, 0
    for r in t.rows:
        if r.kind in ("noise", "page_break"):
            close()
            if r.kind == "noise":
                _noise(cv, r, st)
            continue
        if r.kind == "blank":  # gap above the next Konto line
            gap = 10
            continue
        lines = wrap(r.cells[ni], t.col_widths[ni] - 4, fs) if r.kind in ("asset", "continuation") else [r.cells[ni]]
        h = 14.0 if r.kind == "heading" else 7 + LINE * len(lines)
        if st["y"] is None or st["y"] + gap + h > BOTTOM:
            open_page(first or t.repeat_header)
            first = False
        y, gap = st["y"] + gap, 0
        if r.kind == "heading":
            text(r.cells[0], xs[0] + 2, y + 10, reg if ruled else bold)
            grid(y, y + h, 0, 2)
        else:
            for i, s in enumerate(r.cells):
                for k, s2 in enumerate(lines if i == ni else [s] if s else []):
                    cell(s2, i, y + 10.3 + LINE * k)
            grid(y, y + h)
        st["n"] += 1
        r.pointer = {"page": st["page"], "line": st["n"], "bbox": [xs[0], round(y, 1), xs[-1], round(y + h, 1)]}
        st["y"] = y + h
    close()
    cv.save()


def _noise(cv, r: core.Row, st: dict) -> None:
    """A page that is not part of the register (letter, blank sheet, unrelated list)."""
    W, H = A4
    st["page"] += 1
    cv.setPageSize((W, H))
    cv.setFont("Helvetica", 10)
    for i, s in enumerate(r.cells):
        left, _, right = s.partition("\t")
        cv.drawString(70, H - 90 - i * 14, left)
        cv.drawRightString(480, H - 90 - i * 14, right)
    r.pointer = {"page": st["page"], "line": 1, "bbox": [0, 0, round(W, 1), round(H, 1)]}
    cv.showPage()


def client(seed: int, **kw) -> tuple:
    """Client with free inventory numbers (1-4 digits, some with a letter suffix, 2-digit '01' style for land),
    3-digit accounts, 01.01.YYYY placeholder dates for land. Returns (client, Kostenstelle per class, company block,
    extra invented names)."""
    c = core.make_client(seed, n_classes=13, assets_per_class=(4, 10), **kw)
    rng, used, out = random.Random(f"{FAMILY}-{seed}"), set(), []
    for k, assets in c.by_class():
        n = 0
        for a in assets:
            while True:
                n += 1
                s = f"{n:02d}" if k.kind == "land" else str(rng.choice((rng.randint(1, 99), *[rng.randint(100, 999)] * 2,
                                                                         rng.randint(1000, 4999))))
                if k.kind != "land" and rng.random() < 0.15:
                    s += rng.choice("abcdef" if rng.random() < 0.85 else "ABC")
                if s not in used:
                    break
            used.add(s)
            acq = a.acq_date.replace(month=1, day=1) if k.kind == "land" else a.acq_date
            out.append(replace(a, asset_no=s, account=f"{int(a.account):03d}", acq_date=acq))
    c.assets = out
    kst = {k.code: rng.choice(("0", "0", "")) for k in c.classes}
    street, town = rng.sample([a + b for a in core.SYL_A for b in core.SYL_B], 2)
    street = f"{street}{rng.choice(('straße', 'weg', 'ring'))}"
    town = town[:-2] + rng.choice(("hausen", "stedt", "ingen", "feld"))
    business = rng.choice(("Fertigung und Vertrieb", "Montage und Handel", "Planung und Fertigung", "Werkstatt und Handel"))
    company = [c.company, business, f"{street} {rng.randint(1, 140)}", f"{rng.randint(10000, 99999)} {town}"]
    return c, kst, company, [street, town]


def build(out: Path) -> None:
    d = out / FAMILY
    d.mkdir(parents=True, exist_ok=True)

    def make(seed, cols=COLS, **kw):
        c, kst, company, names = client(seed, year=kw.pop("year", 2025))
        period = f"Anlageverzeichnis vom {core.fmt_date(c.fy_start)} bis {core.fmt_date(c.fy_end)}"
        return c, table(c, cols, kst, period, **kw), company, names

    def key(n, c, names, fmt, variant, t, planted=()):
        meta = core.key_meta(c, doc_id=f"{FAMILY}_{n:02d}", family=FAMILY, doc_type="Anlagenverzeichnis", format=fmt,
                             variant=variant)
        meta["invented_names"] = c.names + names
        core.write_key(d / f"{FAMILY}_{n:02d}.key.json", meta, t, planted)

    def pdf(n, seed, variant, ruled, problems=None, redact=False, **kw):
        c, t, company, names = make(seed, **kw)
        planted = problems(t, c, random.Random(seed)) if problems else []
        render_pdf(t, d / f"{FAMILY}_{n:02d}.pdf", ruled=ruled, company=None if redact else company)
        key(n, c, names, "pdf", variant, t, planted)

    pdf(1, 401, "perfect", ruled=True)
    pdf(2, 402, "perfect", ruled=False, redact=True)
    c, t, company, names = make(403, year=2024)  # 03 the same register as an Excel export
    t.title_lines[1:1] = company
    core.render_xlsx(t, d / f"{FAMILY}_03.xlsx", sheet="Anlageverzeichnis")
    key(3, c, names, "xlsx", "perfect", t)
    for n, seed, fmt, ruled in ((4, 404, "scan", True), (5, 405, "outline", False)):  # no text layer
        c, t, company, names = make(seed)
        with tempfile.TemporaryDirectory() as tmp:
            render_pdf(t, Path(tmp) / "print.pdf", ruled=ruled, company=company if ruled else None)
            if fmt == "scan":  # heavy skew
                core.scan(Path(tmp) / "print.pdf", d / f"{FAMILY}_{n:02d}.pdf", skew_deg=2.5, noise=0.03, seed=seed)
            else:
                core.outline(Path(tmp) / "print.pdf", d / f"{FAMILY}_{n:02d}.pdf")
        key(n, c, names, fmt, "perfect", t)
    # 06: descriptions broken into an extra row + one Konto Summe that does not tie
    pdf(6, 406, "problem", ruled=True, problems=lambda t, c, rng: [core.p_wrap(t, rng, 3)[1], core.p_non_tying(t, rng)[1]])
    # 07: headings only on page 1 + an unrelated print in the middle
    pdf(7, 407, "problem", ruled=False, redact=True, year=2024,
        problems=lambda t, c, rng: [core.p_no_repeat_header(t)[1], core.p_noise(t, c, "unrelated", "middle")[1]])
    # 08: columns reordered and renamed, no thousands dots, Konto line with its name
    pdf(8, 408, "layout-mutation", ruled=False, cols=MUT, num_style="de_nodots", konto="Konto {acc} {name}")
