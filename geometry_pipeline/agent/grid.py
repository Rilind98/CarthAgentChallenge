"""Words with boxes -> raw grid: visual lines x column strips, no meaning attached.

Everything here is layout-agnostic geometry:
  lines  - words grouped by vertical position
  strips - x-ranges covered by amount tokens; a gap bridged by only a few tokens (glued numbers) stays a gap
  cells  - amount tokens of a line assigned to strips; split tokens merged, glued tokens split
"""
import re
from dataclasses import dataclass, field

AMOUNT = re.compile(r"^\(?-?\d[\d.]*(,\d\d)?-?\)?$")
PARTIAL = re.compile(r"^\(?-?\d{1,3}(\.\d{3})*(\.|,\d?)?$")  # an amount cut off at the edge: '71.777', '1.003,'

GLUED = re.compile(r"\(?-?\d[\d.]*?,\d\d-?\)?")
DASH = {"-", "–", "—"}


@dataclass
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Cell:
    text: str
    box: tuple[float, float, float, float]
    partial: bool = False  # cut off: only the leading digits are printed


@dataclass
class Line:
    id: int
    y0: float
    y1: float
    words: list[Word]
    cells: dict[int, Cell] = field(default_factory=dict)
    label: str = ""
    unreadable: set[int] = field(default_factory=set)  # strips with ink we could not read: no value, not zero


@dataclass
class PageGrid:
    page: int
    lines: list[Line]
    strips: list[tuple[float, float]]

    @property
    def amount_lines(self) -> list[Line]:
        return [ln for ln in self.lines if ln.cells]


TEUR_AMOUNT = re.compile(r"^\(?-?(0|[1-9]\d{0,2}(\.\d{3})*)-?\)?$")  # 49  4.303  1.234.567-  (no ids, no dates)


def is_amount(text: str, teur: bool) -> bool:
    if teur:
        return bool(TEUR_AMOUNT.match(text))
    return (bool(AMOUNT.match(text)) and "," in text) or text == "0"  # Excel prints often show zero as '0'


DATE = re.compile(r"^\d{1,2}\.\d{1,2}\.(\d{2}|\d{4})$")
# cents after a dot: a scanned comma that lost its tail (1.234.56 -> 1.234,56). Thousands groups have three
# digits, so a last group of two can only be cents. Read so only inside an amount column (rates, quantities and
# dates like 09.26 stand elsewhere); amounts never start with 0 and another digit.
DOT_CENTS = re.compile(r"^(\(?-?(?!0\d)\d{1,3}(?:\.\d{3})*)\.(\d\d-?\)?)$")


def is_partial(text: str, teur: bool) -> bool:
    return not teur and bool(PARTIAL.match(text)) and not is_amount(text, teur)


def glued_pieces(text: str, teur: bool = False) -> list[str]:
    """Several complete amounts printed without a gap.
    EUR:  '12.977,004.076,11' -> ['12.977,00', '4.076,11']   (each piece ends in ,dd)
    TEUR: '4.3034.446'        -> ['4.303', '4.446']          (only when thousands groups make the cut certain)"""
    if teur:
        if not re.fullmatch(r"[\d.]+-?", text) or TEUR_AMOUNT.match(text):
            return []
        pieces = re.findall(r"[1-9]\d{0,2}(?:\.\d{3})+|0", text)
        return pieces if len(pieces) >= 2 and "".join(pieces) == text else []
    pieces = GLUED.findall(text) if text.count(",") >= 2 else []
    return pieces if len(pieces) >= 2 and "".join(pieces) == text else []


def group_lines(words: list[Word]) -> list[Line]:
    words = sorted(words, key=lambda w: w.yc)
    if not words:
        return []
    h = sorted(w.y1 - w.y0 for w in words)[len(words) // 2]
    lines, cur = [], [words[0]]
    for w in words[1:]:
        if w.yc - cur[-1].yc > 0.5 * h:
            lines.append(cur)
            cur = []
        cur.append(w)
    lines.append(cur)
    return [Line(i, min(w.y0 for w in ln), max(w.y1 for w in ln), sorted(ln, key=lambda w: w.x0))
            for i, ln in enumerate(lines)]


def find_strips(tokens: list[Word], n_lines: int) -> list[tuple[float, float]]:
    """Runs of x covered by at least `need` amount tokens. need > 1 on busy pages, so that a single glued token
    reaching across the gap between two columns does not merge them."""
    if not tokens:
        return []
    need = 2 if n_lines >= 8 else 1
    edges = sorted({round(t.x0, 1) for t in tokens} | {round(t.x1, 1) for t in tokens})
    out = []
    for a, b in zip(edges, edges[1:]):
        mid = (a + b) / 2
        cover = sum(t.x0 <= mid <= t.x1 for t in tokens)
        if cover >= need:
            if out and a - out[-1][1] <= 1.0:
                out[-1] = (out[-1][0], b)
            else:
                out.append((a, b))
    return out


def strip_index(x0: float, x1: float, strips) -> list[int]:
    """Strips a token box overlaps."""
    return [i for i, (a, b) in enumerate(strips) if x0 < b and x1 > a]


def merge_strips(*strip_lists) -> list[tuple[float, float]]:
    """Union of several pages' strips: overlapping strips become one."""
    out = []
    for a, b in sorted(s for sl in strip_lists for s in sl):
        if out and a <= out[-1][1] + 1.0:
            out[-1] = (out[-1][0], max(out[-1][1], b))
        else:
            out.append((a, b))
    return out


def fits(small: list, big: list) -> bool:
    """Same layout, columns empty on one page or the other: no strip of either overlaps two strips of the other,
    and at least half of the smaller page's strips line up with the other's."""
    shared = sum(len(strip_index(a, b, big)) == 1 for a, b in small)
    return (all(len(strip_index(a, b, big)) <= 1 for a, b in small)
            and all(len(strip_index(a, b, small)) <= 1 for a, b in big)
            and shared >= max(1, min(len(small), len(big)) / 2))


def build(page: int, words: list[Word], teur: bool = False, strips: list | None = None) -> PageGrid:
    lines = group_lines(words)
    amount = [w for ln in lines if any(is_amount(w.text, teur) for w in ln.words)
              for w in ln.words if is_amount(w.text, teur) or is_partial(w.text, teur)]
    n_amount_lines = sum(any(is_amount(w.text, teur) for w in ln.words) for ln in lines)
    if strips is None:
        strips = find_strips(amount, n_amount_lines)
    if not teur:
        for ln in lines:
            ln.words = [Word(w.x0, w.y0, w.x1, w.y1, DOT_CENTS.sub(r"\1,\2", w.text))
                        if DOT_CENTS.match(w.text) and strip_index(w.x0, w.x1, strips) else w for w in ln.words]
    for ln in lines:
        has_amount = any(is_amount(w.text, teur) for w in ln.words)
        parts: dict[int, list[Word]] = {}
        label = []
        for w in ln.words:
            hit = strip_index(w.x0, w.x1, strips)
            pieces = glued_pieces(w.text, teur)
            partial = has_amount and is_partial(w.text, teur)
            if has_amount and (is_amount(w.text, teur) or w.text in DASH or pieces or partial) and hit:
                if pieces:  # glued: amounts printed without a gap, aligned at the right-most one's column
                    last = max(hit, key=lambda i: strips[i][1])
                    targets = list(range(last - len(pieces) + 1, last + 1))
                    if targets[0] >= 0:
                        x, per = w.x0, (w.x1 - w.x0) / len(w.text)
                        for i, p in zip(targets, pieces):  # each piece keeps its own share of the box
                            parts.setdefault(i, []).append(Word(x, w.y0, x + per * len(p), w.y1, p))
                            x += per * len(p)
                    else:
                        ln.unreadable.update(hit)
                    continue
                i = min(hit, key=lambda i: abs((strips[i][0] + strips[i][1]) / 2 - w.xc))
                parts.setdefault(i, []).append(w)
            elif has_amount and hit and re.fullmatch(r"[\d.,()\-]+", w.text) and re.search(r"\d", w.text) \
                    and not DATE.match(w.text):
                ln.unreadable.update(hit)  # number-like ink we cannot read: no value rather than a false zero
            elif not (hit and is_amount(w.text, teur)):
                label.append(w.text)
        for i, ws in parts.items():
            ws.sort(key=lambda w: w.x0)
            text = "".join(w.text for w in ws)
            if len(ws) > 1 and not (is_amount(text, teur) or is_partial(text, teur)):
                # not one amount split in pieces: keep the single token that is an amount on its own (a page
                # number or stray mark can share the line), else the cell cannot be read
                whole = [w for w in ws if is_amount(w.text, teur)]
                if len(whole) != 1:
                    ln.unreadable.add(i)
                    continue
                ws, text = whole, whole[0].text
            ln.cells[i] = Cell(text, (min(w.x0 for w in ws), min(w.y0 for w in ws), max(w.x1 for w in ws),
                                      max(w.y1 for w in ws)), partial=is_partial(text, teur))
        ln.label = " ".join(label)
    return PageGrid(page, lines, strips)
