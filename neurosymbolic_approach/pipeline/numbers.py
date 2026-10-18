"""Number and date grammar. The only place where printed text becomes a figure.

Amounts are German-formatted: 1.234,56 · 1234,56 · -1.234,56 · 1.234,56- · (1.234,56); "-" alone is zero.
Field values follow the answer keys: positive cents, except the transfer columns, which keep their sign.
"""

from __future__ import annotations

import re
from datetime import date

_AMT = r"\d{1,3}(?:\.\d{3})+,\d{2}|\d+,\d{2}"
AMOUNT_RE = re.compile(rf"^(?P<lp>\()?(?P<lm>[-−])?(?P<n>{_AMT})(?P<tm>-)?(?P<rp>\))?$")
TEUR_RE = re.compile(r"^(?P<lp>\()?(?P<lm>[-−])?(?P<n>\d{1,3}(?:\.\d{3})+|\d+)(?P<tm>-)?(?P<rp>\))?$")
GLUED_RE = re.compile(rf"^([-−]?(?:{_AMT})-?)([-−]?(?:{_AMT})-?)$")
_AMT_EN = r"\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2}"
AMOUNT_EN_RE = re.compile(rf"^(?P<lp>\()?(?P<lm>[-−])?(?P<n>{_AMT_EN})(?P<tm>-)?(?P<rp>\))?$")
GLUED_EN_RE = re.compile(rf"^([-−]?(?:{_AMT_EN})-?)([-−]?(?:{_AMT_EN})-?)$")
DATE_RE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})$")
DATE_IN_TEXT_RE = re.compile(r"(?<![\d.])(\d{1,2})\.(\d{1,2})\.(\d{4})(?![\d])")
DASHES = {"-", "–", "—", "−"}

SIGNED_FIELDS = {"umbuchung", "bew_ahk"}  # every other field is stored as a positive amount


def detect_style(texts) -> str:
    """'de' (1.234,56) or 'en' (1,234.56): vote over amount-like strings; dates do not vote."""
    de = en = 0
    for t in texts:
        t = t.strip().strip("()-")
        if DATE_RE.match(t) or DATE_IN_TEXT_RE.search(t):
            continue
        if re.fullmatch(_AMT, t):
            de += 1
        elif re.fullmatch(_AMT_EN, t):
            en += 1
    return "en" if en > de else "de"


def parse_amount(text: str, teur: bool = False, style: str = "de") -> int | None:
    """Signed cents, or None when the text is not an amount."""
    t = text.strip()
    if style == "en":
        m = AMOUNT_EN_RE.match(t)
        if m:  # 1,234.56 -> 1234,56, then the German path
            t = t.replace(",", "").replace(".", ",")
    m = AMOUNT_RE.match(t) or (TEUR_RE.match(t) if teur else None)
    if not m:
        return None
    n = m["n"].replace(".", "")
    if "," in n:
        whole, dec = n.split(",")
        cents = int(whole) * 100 + int(dec)
    else:
        cents = int(n) * 100
    if teur:
        cents *= 1000
    neg = bool(m["lm"] or m["tm"] or (m["lp"] and m["rp"]))
    return -cents if neg else cents


def split_glued(text: str, style: str = "de") -> tuple[str, str] | None:
    """'8.421,128.530,49' -> ('8.421,12', '8.530,49'). Unique because amounts carry exactly two decimals."""
    t = text.strip()
    if (AMOUNT_EN_RE if style == "en" else AMOUNT_RE).match(t):
        return None
    m = (GLUED_EN_RE if style == "en" else GLUED_RE).match(t)
    return (m[1], m[2]) if m else None


def field_value(field: str, signed_cents: int) -> int:
    return signed_cents if field in SIGNED_FIELDS else abs(signed_cents)


def parse_date(text: str) -> date | None:
    m = DATE_RE.match(text.strip())
    if not m:
        return None
    d, mo, y = int(m[1]), int(m[2]), m[3]
    year = int(y) if len(y) == 4 else (2000 + int(y) if int(y) < 50 else 1900 + int(y))
    try:
        return date(year, mo, d)
    except ValueError:
        return None


def dates_in(text: str) -> list[date]:
    out = []
    for d, mo, y in DATE_IN_TEXT_RE.findall(text):
        try:
            out.append(date(int(y), int(mo), int(d)))
        except ValueError:
            pass
    return out


def token_kind(text: str, teur: bool = False, style: str = "de") -> str:
    t = text.strip()
    if t in DASHES:
        return "dash"
    if parse_date(t):
        return "date"
    if (AMOUNT_EN_RE if style == "en" else AMOUNT_RE).match(t):
        return "amount"
    if re.fullmatch(r"\d+", t):
        return "int"
    return "text"


def fmt_de(cents: int) -> str:
    s = f"{abs(cents) / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if cents < 0 else "") + s


if __name__ == "__main__":
    assert parse_amount("19.425,96") == 1942596 and parse_amount("2.047,95-") == -204795
    assert parse_amount("(1.234,56)") == -123456 and parse_amount("1234,56") == 123456
    assert parse_amount("49", teur=True) == 4_900_000 and parse_amount("40-", teur=True) == -4_000_000
    assert parse_amount("12.03.2025") is None and parse_amount("173") is None
    assert split_glued("8.421,128.530,49") == ("8.421,12", "8.530,49") and split_glued("0,000,00") == ("0,00", "0,00")
    assert field_value("kum_afa_open", -870463) == 870463 and field_value("umbuchung", -30659684) == -30659684
    assert parse_date("13.11.14") == date(2014, 11, 13) and parse_date("04.05.1988") == date(1988, 5, 4)
    assert fmt_de(-123456) == "-1.234,56"
    assert parse_amount("1,234.56", style="en") == 123456 and parse_amount("(1,234.56)", style="en") == -123456
    assert parse_amount("12.50", style="en") == 1250 and token_kind("31.12.2025", style="en") == "date"
    assert detect_style(["1,234.56", "0.00", "31.12.25", "12.03.2025"]) == "en" and detect_style(["1.234,56", "0,00"]) == "de"
    assert split_glued("8,421.128,530.49", "en") == ("8,421.12", "8,530.49")
    print("numbers OK")
