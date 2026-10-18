from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re

from .models import AmountRepresentation


class NumericParseError(ValueError):
    pass


_SPACE = re.compile(r"[\s\u00a0\u202f]")


def parse_amount(raw: str, *, scale: Decimal = Decimal(1), cr_dr: bool = False) -> AmountRepresentation:
    original = raw
    token = raw.strip()
    if not token or token.upper() in {"N/A", "NA"} or token in {"–", "—", "-"}:
        raise NumericParseError("not a numeric amount")
    negative = False
    notation = "leading_plus_or_unsigned"
    if token.startswith("(") and token.endswith(")"):
        negative, notation, token = True, "parentheses", token[1:-1].strip()
    if token.endswith("-"):
        negative, notation, token = True, "trailing_minus", token[:-1].strip()
    elif token.startswith("-"):
        negative, notation, token = True, "leading_minus", token[1:].strip()
    upper = token.upper()
    if cr_dr and (upper.endswith(" CR") or upper.endswith(" DR")):
        marker = upper[-2:]
        token = token[:-3].strip()
        negative = marker == "CR"
        notation = marker.lower()
    token = _SPACE.sub("", token)
    if not re.fullmatch(r"\d[\d.,]*", token):
        raise NumericParseError(f"unconsumed numeric characters in {original!r}")
    dots, commas = token.count("."), token.count(",")
    decimal_sep: str | None = None
    thousands_sep: str | None = None
    if dots and commas:
        decimal_sep = "." if token.rfind(".") > token.rfind(",") else ","
        thousands_sep = "," if decimal_sep == "." else "."
    elif dots or commas:
        sep = "." if dots else ","
        count = dots or commas
        tail = len(token) - token.rfind(sep) - 1
        if count > 1:
            groups = token.split(sep)
            if not all(len(g) == 3 for g in groups[1:]):
                raise NumericParseError("ambiguous repeated separator")
            thousands_sep = sep
        elif tail in (1, 2):
            decimal_sep = sep
        elif tail == 3:
            thousands_sep = sep
        else:
            raise NumericParseError("ambiguous separator")
    canonical = token
    if thousands_sep:
        canonical = canonical.replace(thousands_sep, "")
    if decimal_sep:
        canonical = canonical.replace(decimal_sep, ".")
    try:
        display = Decimal(canonical)
    except InvalidOperation as exc:
        raise NumericParseError(str(exc)) from exc
    if negative:
        display = -display
    precision = len(canonical.partition(".")[2])
    return AmountRepresentation(original, decimal_sep, thousands_sep, notation, abs(display), display, scale, display * scale, precision)

