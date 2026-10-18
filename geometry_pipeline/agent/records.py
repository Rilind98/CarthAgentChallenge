"""Grid lines -> records under the layout's rules, and the row laws that tie a record's fields together.

Pure code, shared by run.py (writes the output) and ocr_tools.py (checks a scanned page while it is proofread).
"""
import re

from grid import PageGrid
from layout import kind_of, walk

SIGNED = {"umbuchung", "bew_ahk", "bew_wb"}
# row laws on the key convention (positive amounts, signed transfers): target = sum(sign * term)
LAWS = [
    ("ahk_close", {"ahk_open": 1, "zugang": 1, "abgang": -1, "umbuchung": 1}),
    ("ahk_open", {"ahk_close": 1, "zugang": -1, "abgang": 1, "umbuchung": -1}),
    ("kum_afa_close", {"kum_afa_open": 1, "afa_gj": 1, "afa_abgang": -1, "zuschreibung": -1}),
    ("kum_afa_open", {"kum_afa_close": 1, "afa_gj": -1, "afa_abgang": 1, "zuschreibung": 1}),
    ("bw_close", {"ahk_close": 1, "kum_afa_close": -1}),
    ("bw_prev", {"ahk_open": 1, "kum_afa_open": -1}),
]

# the same laws for proofreading, one equation each: sum(coef * field) = 0 over a record, fields the record does
# not have count as zero; an equation applies when its `needs` are in the record and `when` holds
CHECKS = [
    ({"ahk_close": -1, "ahk_open": 1, "zugang": 1, "abgang": -1, "umbuchung": 1, "bew_ahk": 1},
     {"ahk_close", "ahk_open"}, lambda f: True),
    ({"kum_afa_close": -1, "kum_afa_open": 1, "afa_gj": 1, "afa_abgang": -1, "zuschreibung": -1, "bew_wb": 1},
     {"kum_afa_close", "kum_afa_open"}, lambda f: True),
    ({"bw_close": -1, "ahk_close": 1, "kum_afa_close": -1}, {"bw_close", "ahk_close", "kum_afa_close"},
     lambda f: True),
    ({"bw_prev": -1, "ahk_open": 1, "kum_afa_open": -1}, {"bw_prev", "ahk_open", "kum_afa_open"}, lambda f: True),
    # end book value straight from the movements, where end cost or end depreciation is not printed
    ({"bw_close": -1, "ahk_open": 1, "zugang": 1, "abgang": -1, "umbuchung": 1, "bew_ahk": 1, "kum_afa_open": -1,
      "afa_gj": -1, "afa_abgang": 1, "zuschreibung": 1, "bew_wb": -1},
     {"bw_close", "ahk_open", "kum_afa_open"}, lambda f: not {"ahk_close", "kum_afa_close"} <= f),
    # net layouts: book values and movements only
    ({"bw_close": -1, "bw_prev": 1, "zugang": 1, "abgang_bw": -1, "umbuchung": 1, "zuschreibung": 1, "afa_gj": -1},
     {"bw_close", "bw_prev"}, lambda f: "ahk_open" not in f),
]


def rebuild(values: dict[str, int], partial: dict[str, str]) -> dict[str, int]:
    """Fields printed only in part (cut off at the edge): compute them from the row laws, in code, and keep a
    result only if its digits start with the digits that are printed."""
    out, changed = {}, True
    known = {f: (v if f in SIGNED else abs(v)) for f, v in values.items()}
    present = set(values) | set(partial)  # fields this layout prints for the record; others count as zero
    while changed:
        changed = False
        for target, terms in LAWS:
            if target not in partial or target in out:
                continue
            if all(t in known for t in terms if t in present):
                v = sum(sgn * known.get(t, 0) for t, sgn in terms.items())
                if v >= 0 and str(v).startswith(re.sub(r"\D", "", partial[target])):
                    out[target] = known[target] = v
                    changed = True
    return out


def parse_amount(text: str, teur: bool) -> int | None:
    """Printed amount -> cents. 1.234,56  1.234,56-  -1.234,56  (1.234,56)  -  (TEUR: 1.234)"""
    t = text.strip()
    if t in {"-", "–", "—", "", "0"}:
        return 0
    neg = t.startswith("-") or t.endswith("-") or (t.startswith("(") and t.endswith(")"))
    digits = re.sub(r"\D", "", t)
    if not digits:
        return None
    if teur:
        cents = int(digits) * 100_000
    else:
        if not re.search(r",\d\d\)?-?$", t):
            return None
        cents = int(digits)
    return -cents if neg else cents


def rules_start(rules: list[dict]) -> dict | None:
    """The rule whose line opens a record: the first record-line rule the model listed."""
    return next((r for r in rules if r.get("after") is None), None)


def page_records(g: PageGrid, rules: list[dict], kinds: list[dict], headers: dict, teur: bool) -> list[tuple]:
    """Records of one table page in reading order: [(row label, {field: (value in cents, box, printed)})].
    Values are positive except the signed fields; printed = False for a zero taken from an empty cell."""
    context = {ln.id: "" for ln in g.lines}
    cur = ""
    for ln in g.lines:  # the last text line above each line: heading / record key line
        if not ln.cells and ln.label:
            cur = ln.label
        context[ln.id] = cur
    out = []
    last_row = ""
    record: dict[str, tuple] = {}  # field -> (value, box, printed): one value per field and record
    partial: dict[str, tuple] = {}  # field -> (printed fragment, box): cut off, to be rebuilt

    def put(field, v, box, printed: bool):
        old = record.get(field)
        if old is None or (printed and not old[2]):
            record[field] = (v, box, printed)

    def flush(row):
        if partial:
            for field, v in rebuild({f: r[0] for f, r in record.items() if r[2]},
                                    {f: t for f, (t, _) in partial.items()}).items():
                record[field] = (v, partial[field][1], True)
            partial.clear()
        if record:
            out.append((row, {f: (v if f in SIGNED else abs(v), box, printed)
                              for f, (v, box, printed) in record.items()}))
        record.clear()

    def take(ln, r, absent: bool = False):
        """Values of one line under rule r; absent = the continuation line is not printed: its fields are 0."""
        for s, field in r.get("strips", {}).items():
            if field == "ignore":
                continue
            if s.startswith("H"):  # header of a column with no amount anywhere: zero under the header
                if s in headers:
                    a, b = headers[s]
                    put(field, 0, (a, ln.y0, b, ln.y1), False)
                continue
            i = int(s.lstrip("S"))
            if i >= len(g.strips):
                continue
            a, b = g.strips[i]
            if not absent and i in ln.unreadable:
                continue  # ink we could not read: leave the value out, never write 0
            if not absent and i in ln.cells and ln.cells[i].partial:
                partial[field] = (ln.cells[i].text, ln.cells[i].box)
                continue
            if not absent and i in ln.cells:
                v = parse_amount(ln.cells[i].text, teur)
                if v is not None:
                    put(field, v, ln.cells[i].box, True)
            else:
                put(field, 0, (a, ln.y0, b, ln.y1), False)

    for ln, r, nxt in walk(g, rules, kinds):
        if r is None:
            continue
        if r.get("after") is None:
            if r.get("label_regex") is None and r.get("position") is None or not record:
                flush(last_row)  # single-line layouts: every line is a record
            elif r is rules_start(rules):
                flush(last_row)
            if not record:
                last_row = ln.label or context[ln.id]
        take(ln, r)
        if r.get("after") is None:
            for c in rules:
                if c.get("after") == r.get("id") and nxt is not c:
                    take(ln, c, absent=True)
    flush(last_row)
    return out


def broken_laws(fields: dict[str, tuple]) -> list[tuple[dict, int]]:
    """Equations of CHECKS that a record breaks: [(coefs restricted to the record's fields, residual in cents)]."""
    have = set(fields)
    out = []
    for coefs, needs, when in CHECKS:
        if not needs <= have or not when(have):
            continue
        used = {f: c for f, c in coefs.items() if f in have}
        residual = sum(c * fields[f][0] for f, c in used.items())
        if residual:
            out.append((used, residual))
    return out


SUBTOTALS = {"position_subtotal", "account_subtotal", "class_subtotal", "grand_total"}


def broken_sums(recs: list[tuple], kinds: list[dict]) -> list[tuple[int, list[int], dict[str, int]]]:
    """Subtotal rows whose columns are not the sum of the rows they cover: [(subtotal index, span, {field:
    residual})], residual = subtotal - sum of the span.

    Which rows a subtotal covers is not printed, so code tries several spans of asset (and carry-over) rows: since
    the previous subtotal of any kind, since the previous one of the same kind, since the top of the page. It takes
    the span under which most nonzero columns add up and reports the columns that do not, only if at least one
    column adds up (otherwise the subtotal covers rows we cannot see, e.g. on the page before)."""
    kind = [kind_of(row, kinds) for row, _ in recs]
    out = []
    for i, (_, sub) in enumerate(recs):
        if kind[i] not in SUBTOTALS:
            continue
        spans = []
        for stop in (lambda k: k in SUBTOTALS, lambda k: k == kind[i], lambda k: False):
            j = i - 1
            while j >= 0 and not stop(kind[j]):
                j -= 1
            span = [m for m in range(j + 1, i) if kind[m] in ("asset", "carry_over")]
            if span and span not in spans:
                spans.append(span)
        best = None
        for span in spans:
            residual = {f: v - sum(recs[m][1].get(f, (0,))[0] for m in span)
                        for f, (v, _, printed) in sub.items() if printed}
            good = sum(1 for f, r in residual.items() if r == 0 and sub[f][0] != 0)
            if best is None or good > best[0]:
                best = (good, span, residual)
        if best and best[0] >= 1:
            bad = {f: r for f, r in best[2].items() if r}
            if bad:
                out.append((i, best[1], bad))
    return out


def fmt(cents: int, teur: bool) -> str:
    """Cents -> the way the page prints it: 1.234.567,89 (TEUR: 1.235)."""
    if teur:
        return f"{'-' if cents < 0 else ''}{abs(cents) // 100_000:,}".replace(",", ".")
    s = f"{abs(cents) // 100:,}".replace(",", ".") + f",{abs(cents) % 100:02d}"
    return ("-" if cents < 0 else "") + s


def digit_distance(a: str, b: str) -> int:
    """Edits between the digit strings of two amounts (misread, lost or extra digits)."""
    a, b = re.sub(r"\D", "", a), re.sub(r"\D", "", b)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def sum_suspects(recs: list[tuple], kinds: list[dict], teur: bool) -> list[dict]:
    """Cells that could explain a broken column sum (broken_sums over the whole document) with one misread digit:
    the sum cell itself or a cell of the rows it covers. recs: [(row, fields, page)].
    -> [{page, box, text, row, field, note}]"""
    out = []
    for i, span, bad in broken_sums([(r, f) for r, f, _ in recs], kinds):
        row, sub, page = recs[i]
        for field, residual in bad.items():
            note = f"sum row '{row[:40]}' (page {page}) is off by {fmt(residual, teur)} in this column"
            cells = [(i, sub[field][0] - residual)] + [(m, recs[m][1][field][0] + residual) for m in span
                                                       if field in recs[m][1] and recs[m][1][field][2]]
            for m, need in cells:
                v, box, _ = recs[m][1][field]
                if v and need and (need > 0 or field in SIGNED) \
                        and digit_distance(fmt(v, teur), fmt(need, teur)) <= 1:
                    out.append({"page": recs[m][2], "box": box, "text": fmt(v, teur), "row": recs[m][0],
                                "field": field, "note": note + f"; {fmt(need, teur)} here would fix it"})
    return out
