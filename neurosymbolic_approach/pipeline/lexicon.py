"""Domain knowledge shared by every layout: header vocabulary and the accounting laws.

A field is GROUP x ROLE. Group: AHK (acquisition cost), AFA (depreciation), BW (book value).
Role: open, add, disposal, transfer, writeup, year, close. So "Zugänge" under "kumulierte Abschreibungen" is afa_gj,
and "Zugänge" with no group is zugang. Nothing here names a layout or a family.
"""

from __future__ import annotations

import re
from datetime import date, timedelta
from fractions import Fraction

from .numbers import dates_in

GROUPS = {
    "AHK": re.compile(r"\bahk\b|anschaff|herstell|\bansch\b|ansch[-./]|anschwert|\bak\b"),
    "AFA": re.compile(r"\bafa\b|abschreib|\babsch\b|wertber|\babschr\b"),
    "BW": re.compile(r"buchw|buchwrt|\bbw\b"),
}
HARD_EXCLUDE = re.compile(r"zinsen|sonder|nachaktiv|invest|f[oö]rder|lt\.?\s*bilanz|differenz|kosten-?\s*stelle|"
                          r"w[aä]hr|datum|summe|afa[- ]?art|method|anzahl|%|satz|nutzungsdauer|kostenst|\bkst\b")
EXCLUDE = re.compile(r"\bnd\b|\bart\b|bilpos|klasse|konto|bezeichnung|inventar|anlage\b|unr\b|unter|aktiv")
# labels of non-amount columns (dates, currency, method, life, cost centre, ids); amount columns we leave unmapped
# (Zinsen, Sonder-AfA, Differenz, Summe) are NOT text columns
TEXT_COLUMN = re.compile(r"datum|w[aä]hr|method|nutzungsdauer|kostenst|\bkst\b|afa[- ]?art|anzahl|%|satz|"
                         r"\bnd\b|\bart\b|bilpos|klasse|konto|bezeichnung|inventar|anlage\b|unr\b|unter|aktiv")
NET = [("bew_ahk", re.compile(r"\bbew(\.|egung(en)?)?\s*(ansch|ahk)")),
       ("bew_wb", re.compile(r"\bbew(\.|egung(en)?)?\s*(wertber|wb\b|afa|abschr)"))]
ROLES = [  # strongest first
    ("transfer", re.compile(r"umbuch")),
    ("writeup", re.compile(r"zuschreib")),
    ("disposal", re.compile(r"abg[aä]ng|minderung")),
    ("add", re.compile(r"zug[aä]ng")),
    ("year", re.compile(r"des\s+jahres|lfd\.?\s*jahr|normal-?afa|gep\.?\s*afa|afa\s+jahr|planm")),
    ("open", re.compile(r"beginn|gj-?\s?beg|\bbeg\b|beg-|anfang|vorjahr|\bhist\b|hist\.")),
    ("close", re.compile(r"\bende\b|ende-|aktuell|\blfd\b|endstand|kumuliert|31\.12\.?(?!\d)")),
]
WEAK_ROLES = [("year", re.compile(r"gesch[aä]ftsjahr|\bgj\b")), ("kum", re.compile(r"\bkum\b|kum\."))]
TABLE = {
    "AHK": {"open": "ahk_open", "close": "ahk_close", "add": "zugang", "disposal": "abgang", "transfer": "umbuchung",
            "writeup": "zuschreibung"},
    "AFA": {"open": "kum_afa_open", "close": "kum_afa_close", "add": "afa_gj", "year": "afa_gj",
            "disposal": "afa_abgang", "writeup": "zuschreibung"},
    "BW": {"open": "bw_prev", "close": "bw_close", "disposal": "abgang_bw"},
    None: {"add": "zugang", "disposal": "abgang", "transfer": "umbuchung", "writeup": "zuschreibung", "year": "afa_gj"},
}
ID_COLUMNS = [  # identity columns (xlsx headers)
    ("sub_no", re.compile(r"\bunr\b|unr\.|unter-?\s*nummer|unternr")),
    ("asset_no", re.compile(r"^anlage$|^anlage\s*$|inventar-?\s*nummer|inv\.?-?\s*nr|anlagennummer|^anlage\b(?!n)")),
    ("acq_date", re.compile(r"aktiv|anschaff\.?\s*datum|akt\.-?datum|zugangsdatum")),
    ("name", re.compile(r"bezeichnung")),
    ("class", re.compile(r"klasse")),
    ("position", re.compile(r"bilpos|bilanzpos")),
    ("account", re.compile(r"konto")),
    ("method", re.compile(r"afa[- ]?art")),
]


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\n", " ")).strip().lower()


def candidates(header: str, group_header: str = "", period_start: date | None = None,
               period_end: date | None = None) -> list[tuple[str, float, str]]:
    """Ranked (field, score, reason) for one column. Empty list = not a schedule amount."""
    h, g = norm(header), norm(group_header)
    for field, rx in NET:
        if rx.search(h):
            return [(field, 1.0, f"net column '{header}'")]
    if not h and not g:
        return []
    if HARD_EXCLUDE.search(h):
        return []
    if EXCLUDE.search(h) and not any(rx.search(h) for _, rx in ROLES[:4]):
        return []
    if re.search(r"afa\s*umbuch|afa\s*nachakt", h):
        return []
    groups = [k for k, rx in GROUPS.items() if rx.search(h)] or [k for k, rx in GROUPS.items() if rx.search(g)]
    roles = [r for r, rx in ROLES if rx.search(h)]
    # a date in the header decides open/close against the period
    for d in dates_in(header) + dates_in(group_header):
        if period_start and d <= period_start:
            roles.append("open")
        elif period_end and d == period_end:
            roles.append("close")
    if not roles:
        roles = [r for r, rx in WEAK_ROLES if rx.search(h)]
    out: dict[str, tuple[float, str]] = {}

    def add(field, score, why):
        if field and score > out.get(field, (0, ""))[0]:
            out[field] = (score, why)

    role = next((r for r, _ in ROLES if r in roles), None) or (roles[0] if roles else None)
    for grp in groups or [None]:
        share = 1.0 / max(1, len(groups))
        if role == "kum":  # "kum AnschWert", "kum AfA": cumulative, at start or end
            add(TABLE.get(grp, {}).get("close"), 0.55 * share, "kum -> close?")
            add(TABLE.get(grp, {}).get("open"), 0.45 * share, "kum -> open?")
        elif role:
            add(TABLE.get(grp, {}).get(role), 1.0 * share, f"{grp or '-'} x {role}")
            if role == "disposal" and grp is None:
                add("abgang_bw", 0.5, "disposal without group: at cost or at book value")
            if role == "open" and grp is None:
                for f in ("ahk_open", "bw_prev", "kum_afa_open"):
                    add(f, 0.3, "open without group")
            if role == "close" and grp is None:
                for f in ("ahk_close", "bw_close", "kum_afa_close"):
                    add(f, 0.3, "close without group")
        elif grp:  # group without role: laws decide; prior = closing value (schedules print the period end)
            add(TABLE[grp]["close"], 0.55 * share, f"{grp} without role -> close?")
            add(TABLE[grp]["open"], 0.45 * share, f"{grp} without role -> open?")
            if grp == "AFA":
                add("afa_gj", 0.3, "AFA without role")
    return sorted(((f, round(s, 3), why) for f, (s, why) in out.items()), key=lambda x: -x[1])


def id_column(header: str) -> str | None:
    h = norm(header)
    for name, rx in ID_COLUMNS:
        if rx.search(h):
            return name
    return None


def period_from(text_lines: list[str]) -> date | None:
    """The period end named in the title lines; print/creation dates are ignored."""
    best = None
    for line in text_lines:
        low = line.lower()
        if re.search(r"erstell|druck|(?<![a-z])datum:|bearbeit|seite|blatt", low):
            continue
        for d in dates_in(line):
            best = d if best is None or d > best else best
    return best


def period_start_of(end: date) -> date:
    try:
        return end.replace(year=end.year - 1) + timedelta(days=1)
    except ValueError:
        return end.replace(year=end.year - 1, day=28) + timedelta(days=1)


# ---------------------------------------------------------------- accounting laws (positive-cents convention)

VARS = ["ahk_open", "zugang", "abgang", "umbuchung", "zuschreibung", "ahk_close", "kum_afa_open", "afa_gj",
        "afa_abgang", "kum_afa_close", "bw_prev", "bw_close", "abgang_bw", "bew_ahk", "bew_wb"]
LAWS = {
    "ahk_rollforward": {"ahk_close": 1, "ahk_open": -1, "zugang": -1, "abgang": 1, "umbuchung": -1},
    "afa_rollforward": {"kum_afa_close": 1, "kum_afa_open": -1, "afa_gj": -1, "afa_abgang": 1, "zuschreibung": 1},
    "bw_close": {"bw_close": 1, "ahk_close": -1, "kum_afa_close": 1},
    "bw_prev": {"bw_prev": 1, "ahk_open": -1, "kum_afa_open": 1},
    "abgang_bw": {"abgang_bw": 1, "abgang": -1, "afa_abgang": 1},
    "bew_ahk": {"bew_ahk": 1, "zugang": -1, "abgang": 1, "umbuchung": -1},
    "bew_wb": {"bew_wb": 1, "afa_abgang": -1, "zuschreibung": -1},
}


def checkable_laws(observed: set[str]) -> list[dict[str, int]]:
    """Linear combinations of LAWS that use only observed fields (Gaussian elimination of the unobserved ones)."""
    unobs = [v for v in VARS if v not in observed]
    order = unobs + [v for v in VARS if v in observed]
    rows = [[Fraction(law.get(v, 0)) for v in order] for law in LAWS.values()]
    r = 0
    for c in range(len(unobs)):
        piv = next((i for i in range(r, len(rows)) if rows[i][c] != 0), None)
        if piv is None:
            continue
        rows[r], rows[piv] = rows[piv], rows[r]
        for i in range(len(rows)):
            if i != r and rows[i][c] != 0:
                f = rows[i][c] / rows[r][c]
                rows[i] = [a - f * b for a, b in zip(rows[i], rows[r])]
        r += 1
    out = []
    for row in rows[r:]:
        coeffs = {order[j]: row[j] for j in range(len(order)) if row[j] != 0}
        if len(coeffs) >= 2 and all(v in observed for v in coeffs):
            den = 1
            for v in coeffs.values():
                den = den * v.denominator // __import__("math").gcd(den, v.denominator)
            out.append({k: int(v * den) for k, v in coeffs.items()})
    return out


def law_name(law: dict[str, int]) -> str:
    return " ".join(f"{'+' if c > 0 else '-'}{abs(c) if abs(c) != 1 else ''}{f}" for f, c in law.items()) + " = 0"


if __name__ == "__main__":
    print(candidates("Zugänge", "kumulierte Abschreibungen"))
    print(candidates("Abgänge / Minderungen Geschäftsjahr (-)EUR"))
    print(candidates("kum AnschWert"), candidates("Restbuchwert 31.12.2024", "", date(2025, 1, 1), date(2025, 12, 31)))
    print(candidates("davon aktivierte Zinsen Fremdkapital"), candidates("AfA lfd. Jahr"), candidates("lfd Buchwert"))
    f01 = {"bw_prev", "zugang", "abgang_bw", "umbuchung", "zuschreibung", "afa_gj", "kum_afa_close", "bw_close"}
    print([law_name(l) for l in checkable_laws(f01)])
    print([law_name(l) for l in checkable_laws({"ahk_close", "bw_close"})])
