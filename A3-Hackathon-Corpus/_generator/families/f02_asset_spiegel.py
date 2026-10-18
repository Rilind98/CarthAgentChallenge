"""f02_asset_spiegel: gross Anlagenspiegel ("Anlagespiegel ... steuerrechtlich") one line per asset, the sister print
of f01_net_spiegel from the same German mid-market fixed-asset module (same client tree, same run).

Layout: landscape, Arial 7 pt, asset lines regular, headings and sums bold, a blank line after every line. Four
levels: numbered Bilanzposition > Gruppe (indent 1) > Konto (indent 2, "<no> <name>"; missing for the groups GWG and
Anlagen im Bau, whose assets hang directly under the group) > Anlage. Asset line: short inventory number (1-4 digits,
letter suffix a, b, … for split items) right-aligned in a narrow column with the description right after it, so the
text layer reads "<id><Description>"; acquisition date DD.MM.YY; AfA Art lin / geb / sb / k; ND in whole years.
Movement columns blank on asset lines; AfA GJ, Buchwert Vorjahr, kum. AfA and Buchwert 31.12. always printed. Sum lines
"Summe <no> <Konto>", "Summe <Gruppe>", "Summe <n>. <Position>" print every column ("0,00"). Disposals at cost. No grand
total. Heading block + unit row with sign hints on every page; footer "Datum:DD.MM.YY" left, "Seite: n" right.

Key: asset_no = inventory number, sub_no = letter suffix ('' if none); class = printed group, position = printed
position, hgb_position = HGB § 266 code, account = printed Konto (None where the Konto level is missing).
"""

import random
import sys
import tempfile
from pathlib import Path

from reportlab.pdfbase.pdfmetrics import stringWidth

sys.path[:0] = [str(Path(__file__).resolve().parents[1]), str(Path(__file__).resolve().parent)]
import core  # noqa: E402
import f01_net_spiegel as f01  # noqa: E402  (shared client tree and print helpers of the same program)

FAMILY = Path(__file__).stem
FIELDS = ["asset_no", "name", "acq_date", None, None, "ahk_open", "zugang", None, "abgang", "umbuchung",
          "zuschreibung", "afa_gj", "bw_prev", "kum_afa_close", "bw_close"]
HEADS = [("Inventar-", "Nummer"), ("Inventar-", "Bezeichnung"), ("Anschaff.", "Datum"), ("AfA", "Art"), ("ND",),
         ("Anschaff.-/", "Herstellungs-", "kosten (hist.)"), ("Zugänge,", "Geschäftsjahr"),
         ("davon aktivierte", "Zinsen", "Fremdkapital"), ("Abgänge /", "Minderungen", "Geschäftsjahr"),
         ("Umbuchungen", "Geschäftsjahr"), ("Zuschreibungen", "Geschäftsjahr"), ("Abschreibungen", "Geschäftsjahr"),
         ("Buchwert", "Vorjahr"), ("Abschreibungen", "(kumuliert)"), ("Buchwert", "31.12.")]
UNITS = ["", "", "", "", "", "EUR", "(+)EUR", "(+)EUR", "(-)EUR", "(+/-)EUR", "(+)EUR", "EUR", "EUR", "(-)EUR", "EUR"]
MOVES = ("zugang", "abgang", "umbuchung", "zuschreibung")  # blank on asset lines when zero
ZERO = (7,)  # aktivierte Zinsen: no key field, '0,00' on sum lines
NO_KONTO = {"GWG", "Anlagen im Bau"}


def art(a) -> str:
    """AfA Art code of the program."""
    if a.method == "linear":
        return "geb" if a.position == "A.II.1" else "lin"
    return "sb" if a.method == "GWG" else "k"


def table(c, seed: int, pdf: bool = True) -> core.Table:
    widths = [30, 150, 32, 16, 14] + [55] * 10
    align, font = ["l", "l", "l", "l", "r"] + ["r"] * 10, "Arial"
    title = f"Anlagespiegel {core.fmt_date(c.fy_end)} steuerrechtlich"
    heads = HEADS[:-1] + [("Buchwert", f"{c.fy_end:%d.%m.}")]
    t = core.Table(fields=FIELDS, col_widths=widths, align=align, zero=None, date_style="de_short", font=font,
                   rules=False, header_rows=f01.header(heads, UNITS, widths, align, "Arial-Bold", pdf),
                   title_lines=[title] + [""] * 5 if pdf else [title, c.company],
                   page_header=c.company if pdf else "", print_date=f01.print_date(c, seed),
                   page_footer=f01.footer("Datum:{date}", "Seite: {p}", sum(widths), font))
    ids = f01.short_ids(c, seed)

    def line(kind, values=None):
        r = t.add(kind, values or {})
        if pdf:
            t.add("blank", {})
        return r

    def summe(kind, label, assets, **key):
        r = t.total(kind, "", assets, **key)
        r.cells[0] = label  # sum labels run from the left edge, like the headings
        f01.zero_cols(r, ZERO)
        t.rows.append(r)
        if pdf:
            t.add("blank", {})

    for pos, groups in f01.tree(c):
        line("heading", {0: pos})
        p_assets = []
        for g, accts in groups:
            line("heading", {0: f01.IND1 + g})
            konto, g_assets = g not in NO_KONTO, []
            for k, assets in accts:
                if konto:
                    line("heading", {0: f"{f01.IND2}{k.account}    {k.name}"})
                for a in assets:
                    r = t.add_asset(a)
                    no, suffix = ids(a)
                    r.cells[0], r.cells[3], r.cells[4] = no + suffix, art(a), str(a.nd_months // 12)
                    for f in MOVES:
                        if not r.truth["fields"][f]:
                            r.cells[FIELDS.index(f)] = r.truth["printed"][f] = ""
                    r.truth.update({"asset_no": no, "sub_no": suffix, "hgb_position": a.position, "position": pos,
                                    "class": g, "account": a.account if konto else None})
                    if pdf:
                        t.add("blank", {})
                if konto:
                    summe("account_subtotal", f"{f01.IND2}Summe {k.account}    {k.name}", assets, account=k.account)
                g_assets += assets
            summe("class_subtotal", f"{f01.IND1}Summe {g}", g_assets, cls=g)
            p_assets += g_assets
        summe("position_subtotal", f"Summe {pos}", p_assets, position=pos)
        if pdf:
            t.add("blank", {})
    return t


def glue(t: core.Table) -> core.Table:
    """The program's text layer: the inventory number right-aligned in its narrow column and the description starting
    right after it, so text extraction reads '<id><Description>'. Run on the PDF table after the problems."""
    sp = stringWidth(" ", t.font, t.font_size)
    for r in t.rows:
        if r.kind == "asset":
            gap = t.col_widths[0] - 3 - stringWidth(r.cells[0], t.font, t.font_size)
            r.cells[0], r.cells[1] = " " * max(0, round(gap / sp)) + r.cells[0] + r.cells[1], ""
    return t


MUT_FIELDS = ["account", "asset_no", "sub_no", "name", "acq_date", None, None, "ahk_open", "zugang", "abgang",
              "umbuchung", "ahk_close", "afa_gj", "kum_afa_close", "bw_close", "bw_prev"]


def table_mutation(c, seed: int) -> core.Table:
    """Same document, other print set-up: one flat list sorted by account (no position / group levels), Konto and
    sub-number in columns of their own, 4-digit years, method spelled out, AHK at year end added, Zuschreibungen and
    Zinsen dropped, Abgang and negatives in parentheses, zero as '-', Arial Narrow, rules, sums per account + total."""
    d0, d1, prev = (core.fmt_date(x) for x in (c.fy_start, c.fy_end, c.fy_end.replace(year=c.year - 1)))
    widths = [26, 24, 14, 140, 38, 14, 30] + [54] * 9
    align = ["l", "r", "l", "l", "l", "r", "l"] + ["r"] * 9
    heads = [("Konto",), ("Inv.-Nr.",), ("UNr",), ("Bezeichnung",), ("Akt.-Datum",), ("ND",), ("Methode",),
             ("AHK", d0), ("Zugang",), ("Abgang",), ("Umbuchung",), ("AHK", d1), ("AfA", "lfd. Jahr"),
             ("AfA kum.", d1), ("Restbuchwert", d1), ("Restbuchwert", prev)]
    t = core.Table(fields=MUT_FIELDS, col_widths=widths, align=align, zero="-", num_style="paren", font="Narrow",
                   header_rows=f01.header(heads, [""] * 16, widths, align, "Narrow-Bold")[:-2],
                   title_lines=[c.company, f"Anlagenspiegel Einzelnachweis Geschäftsjahr {d0} - {d1}",
                                "Bewertungsbereich Steuerbilanz          Beträge in EUR"],
                   page_header="Seite {p} / {n}", page_footer="Druckdatum {date}", print_date=f01.print_date(c, seed))
    ids = f01.short_ids(c, seed)
    order = sorted(c.assets, key=lambda a: (a.account, int(ids(a)[0]), ids(a)[1]))
    for acct in sorted({a.account for a in order}):
        assets = [a for a in order if a.account == acct]
        for a in assets:
            r = t.add_asset(a)
            no, suffix = ids(a)
            r.cells[1:3], r.cells[5], r.cells[6] = [no, suffix], str(a.nd_months // 12), a.method
            if a.abgang:  # outflow printed negative; the key keeps the positive amount
                r.cells[MUT_FIELDS.index("abgang")] = r.truth["printed"]["abgang"] = core.fmt(-a.abgang, "paren")
            r.truth.update({"asset_no": no, "sub_no": suffix, "hgb_position": a.position, "position": None,
                            "class": None})
        r = t.add_total("account_subtotal", "", assets, account=acct)
        r.cells[0] = f"Summe Konto {acct}"
        if v := r.truth["fields"]["abgang"]:
            r.cells[MUT_FIELDS.index("abgang")] = r.truth["printed"]["abgang"] = core.fmt(-v, "paren")
        t.add("blank", {})
    r = t.add_total("grand_total", "", c.assets)
    r.cells[0] = "Summe Anlagevermögen"
    if v := r.truth["fields"]["abgang"]:
        r.cells[MUT_FIELDS.index("abgang")] = r.truth["printed"]["abgang"] = core.fmt(-v, "paren")
    return t


def build(out: Path) -> None:
    d = out / FAMILY
    d.mkdir(parents=True, exist_ok=True)

    def key(n, c, fmt, variant, t, planted=(), **extra):
        meta = core.key_meta(c, doc_id=f"{FAMILY}_{n:02d}", family=FAMILY, doc_type="Anlagenspiegel (brutto, je Anlage)",
                             format=fmt, variant=variant, **extra)
        core.write_key(d / f"{FAMILY}_{n:02d}.key.json", meta, t, planted)

    def pdf(n):
        return d / f"{FAMILY}_{n:02d}.pdf"

    pair = {"paired_with": f01.PAIR_DOCS[0],
            "pairing": "same client and software run as the net account-level print; group (class) and position "
                       "sums agree"}

    c = f01.client(f01.PAIR_SEED)  # 02 perfect text PDF, paired with f01_net_spiegel_03
    t = glue(table(c, f01.PAIR_SEED))
    core.render_pdf(t, pdf(2))
    key(2, c, "pdf", "perfect", t, **pair)

    c = f01.client(202)  # 03 perfect Excel export
    t = table(c, 202, pdf=False)
    core.render_xlsx(t, d / f"{FAMILY}_03.xlsx", sheet="Anlagespiegel")
    f01.xlsx_fix(d / f"{FAMILY}_03.xlsx", t, int_cols=(4,))
    key(3, c, "xlsx", "perfect", t)

    with tempfile.TemporaryDirectory() as tmp:
        c = f01.client(203)  # 06 fax of the print
        t = glue(table(c, 203))
        core.render_pdf(t, Path(tmp) / "print.pdf")
        core.scan(Path(tmp) / "print.pdf", pdf(6), fax=True, skew_deg=0.8, seed=203)
        key(6, c, "scan", "perfect", t)

        c = f01.client(204)  # 04 clean print without text layer
        t = glue(table(c, 204))
        core.render_pdf(t, Path(tmp) / "print.pdf")
        core.outline(Path(tmp) / "print.pdf", pdf(4))
        key(4, c, "outline", "perfect", t)

    c, rng = f01.client(205), random.Random(205)  # 01 problems: wrapped descriptions + rounding drift in sum lines
    t = table(c, 205)
    planted = [core.p_wrap(t, rng, n=10)[1], f01.cent_drift(t, rng, 0.3)[1]]
    core.render_pdf(glue(t), pdf(1))
    key(1, c, "pdf", "problem", t, planted)

    c, rng = f01.client(206), random.Random(206)  # 07 problems: glued amounts, a sum that does not tie, blank page
    t = table(c, 206)
    planted = [core.p_glued(t, rng, n=3)[1], core.p_non_tying(t, rng)[1], core.p_noise(t, c, "blank", "middle")[1]]
    core.render_pdf(glue(t), pdf(7))
    key(7, c, "pdf", "problem", t, planted)

    c = f01.client(208)  # 05 layout mutation
    t = table_mutation(c, 208)
    core.render_pdf(t, pdf(5))
    key(5, c, "pdf", "layout-mutation", t)
