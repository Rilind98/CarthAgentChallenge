"""The one OpenAI model call per document: name the strips and lines of the raw grid. The model never returns an amount.

Input: per group of pages that share the same strips, the strips with the header words above them and sample lines;
plus the first lines of every page. Output (JSON):
  unit         EUR | TEUR
  pages        [{page, role: table | other, year}]
  groups       [{group, rules: [{label_regex, position, strips: {S<i>: field | "ignore"}}]}]
               a line takes the first rule whose label_regex matches its label (null = any) and, if position is
               set, that is the n-th amount line after the last line without amounts (block layouts)
  kinds        [{label_regex, kind}]  first match wins, default "asset"
"""
import json
import os
import re
from pathlib import Path

from grid import PageGrid, fits, merge_strips, strip_index

FIELDS = {
    "ahk_open": "acquisition/production cost (AHK, Anschaffungswert) at the START of the fiscal year",
    "zugang": "additions (Zugang/Zugänge) to cost during the year",
    "abgang": "disposals (Abgang/Abgänge) at cost",
    "umbuchung": "transfers (Umbuchung), signed",
    "zuschreibung": "write-ups / reversals of depreciation (Zuschreibung)",
    "ahk_close": "acquisition/production cost at the END of the fiscal year (aktuelle AHK)",
    "kum_afa_open": "accumulated depreciation at the START of the year (kum. AfA / Wertberichtigung Beginn)",
    "afa_gj": "depreciation of the fiscal year (AfA des Jahres / Geschäftsjahr)",
    "afa_abgang": "accumulated depreciation released on disposals (AfA Abgang)",
    "kum_afa_close": "accumulated depreciation at the END of the year",
    "bw_close": "book value at the END of the fiscal year (Buchwert, lfd. Buchwert)",
    "bw_prev": "book value at the end of the PRIOR year / start of this year (Buchwert Vorjahr, GJ-Beginn)",
    "abgang_bw": "book value of disposals (Restbuchwert der Abgänge), e.g. in net layouts or as the disposals figure on a book value line",
    "bew_ahk": "one net movement column for cost (e.g. 'Bew AnschWert' = additions - disposals + transfers)",
    "bew_wb": "one net movement column for accumulated depreciation (e.g. 'Bew Wertber')",
    "ignore": "anything else: quantities, rates, control/check columns, columns not listed above",
}
KINDS = ["asset", "class_subtotal", "account_subtotal", "position_subtotal", "grand_total", "carry_over", "ignore"]

PROMPT = """You read German fixed-asset schedules (Anlagenspiegel / Anlagengitter / Anlagenverzeichnis).
Code has already cut each page into visual lines and column strips S0, S1, ... (x-ranges where amounts stand).
Your job is ONLY to say what the strips and lines mean. Never output an amount.

Fields a strip can be:
{fields}

Rules:
- Map every strip that holds amounts on table lines. Use the header words above each strip, the page images and
  the arithmetic you can see (e.g. end = start + additions - disposals).
- If one record is printed over several lines (e.g. an AHK line, an AfA line and a book value line, or a label like
  'Ansch-/Herst-K' / 'Abschreibung' / 'Buchwerte' per line), give one rule per line type, each with an "id":
  match it with label_regex (a Python regex searched in the line label) or with position (1 = first amount line
  after the last line without amounts).
- An extra line that belongs to the line above (e.g. negative movements printed on a second line under their
  column) gets its own rule with "after": "<id of the rule of the line above>" and label_regex "^$"; its strips
  say what a figure in that column means on this extra line (e.g. disposals instead of additions).
- Each field is printed once per record. When a line repeats a figure that another line of the same record
  already shows (e.g. a book value line repeating additions, transfers or depreciation), map the repeat to
  "ignore".
- Each field at most once per rule. A column that has no field of its own is "ignore": e.g. special or
  unplanned depreciation printed next to the normal one (Sonder-AfA, außerplanmäßig, 'davon' notes),
  capitalised interest, quantities, rates, check columns.
- A column whose header is printed but that holds no amount on any line still counts: map its header word
  (H id) to the field, in the same strips object, e.g. "H3": "zuschreibung". Empty cells are zero.
- If every line of a group means the same, give a single rule with label_regex null and position null.
- kinds: regexes on the line label that mark subtotals / totals / carry-over lines (e.g. '^Summe', '^\\*{{2}}',
  '^Übertrag'); everything else is an asset line. Use the most specific kind you can tell: class_subtotal,
  account_subtotal, position_subtotal, grand_total, carry_over. Lines that are not schedule rows at all but still
  carry a number (page headers with report/print dates, page numbers, selection criteria, a repeated column
  header line, reconciliation / check blocks below the table such as 'Abstimmung', '... lt. Bilanz', 'Differenz')
  get kind "ignore".
- pages: role "table" if the page holds schedule lines, "other" for covers, letters, blank or unrelated pages.
  year = the fiscal year the page's table belongs to (the year of its closing date).
- unit: "TEUR" if amounts are in thousand euro, else "EUR".

Answer with JSON only:
{{"unit": "EUR", "pages": [{{"page": 1, "role": "table", "year": 2025}}],
  "groups": [{{"group": 0, "rules": [{{"id": "main", "label_regex": null, "position": null, "after": null,
                                     "strips": {{"S0": "ahk_open"}}}}]}}],
  "kinds": [{{"label_regex": "^Summe", "kind": "class_subtotal"}}]}}

{doc}"""


def group_pages(grids: list[PageGrid], separate: bool = False) -> list[dict]:
    """Pages whose strips fit into another page's strips share one layout; the group keeps the union of strips,
    so a column that is empty on one page still exists there. separate: every page its own group (spreadsheet
    sheets sit on one fixed cell grid, so their strips always fit even when the columns are shifted)."""
    groups: list[dict] = []
    for g in sorted((g for g in grids if g.strips), key=lambda g: -len(g.strips)):
        home = None if separate else next((gr for gr in groups if fits(g.strips, gr["strips"])), None)
        if home:
            home["pages"].append(g.page)
            home["strips"] = merge_strips(home["strips"], g.strips)
        else:
            groups.append({"pages": [g.page], "strips": list(g.strips)})
    for gr in groups:
        gr["pages"].sort()
    return groups


def header_words(g: PageGrid, fallback: PageGrid | None) -> tuple[dict[int, list[str]], list]:
    """Words of the lines above the first amount line: those over a strip, and the others (possible empty
    columns) as (id, text, x0, x1)."""
    over: dict[int, list[str]] = {}
    loose = []
    for src in (g, fallback):
        if src is None:
            continue
        first = next((ln.id for ln in src.lines if ln.cells), None)
        for ln in src.lines:
            if first is None or ln.id >= first:
                break
            for w in ln.words:
                hit = strip_index(w.x0 - 4, w.x1 + 4, g.strips)
                for i in hit:
                    over.setdefault(i, []).append(w.text)
                if not hit and g.strips and w.x0 > g.strips[0][0] - 60:
                    loose.append(w)
        if over:
            break
    clusters: list[list] = []  # words stacked over one x-range form one header
    for w in sorted(loose, key=lambda w: w.x0):
        if clusters and w.x0 <= max(v.x1 for v in clusters[-1]) + 3:
            clusters[-1].append(w)
        else:
            clusters.append([w])
    free = []
    for cl in clusters:
        text = " ".join(v.text for v in sorted(cl, key=lambda v: (round(v.y0), v.x0)))
        if re.fullmatch(r"(\(?[+\-/]*\)?\s*T?EUR\s*)+", text):
            continue  # a unit label alone is not a column header
        free.append((f"H{len(free)}", text, min(v.x0 for v in cl), max(v.x1 for v in cl)))
    return over, free


def describe(grids: list[PageGrid], groups: list[dict], max_lines: int = 40) -> str:
    parts = ["PAGES (first text lines):"]
    for g in grids:
        heads = [" ".join(w.text for w in ln.words) for ln in g.lines[:4]]
        parts.append(f"  page {g.page}: " + " | ".join(h[:90] for h in heads))
    by_page = {g.page: g for g in grids}
    first_table = next((g for g in grids if g.strips), None)
    for gi, gr in enumerate(groups):
        rep = max((by_page[p] for p in gr["pages"]), key=lambda g: len(g.amount_lines))
        hw, free = header_words(rep, first_table)
        gr["headers"] = {h: (x0, x1) for h, _, x0, x1 in free}
        parts.append(f"\nGROUP {gi}: pages {gr['pages']} (representative page {rep.page})")
        for i, (a, b) in enumerate(rep.strips):
            parts.append(f"  S{i}: x {a:.0f}-{b:.0f}   header words above: {' '.join(hw.get(i, [])) or '-'}")
        if free:
            parts.append("  header words over no strip (a column that is empty on every line; map it with its "
                         "H id like a strip if it is a field): "
                         + "  ".join(f"{h}={t}@x{x0:.0f}-{x1:.0f}" for h, t, x0, x1 in free))
        parts.append("  lines (id | label | cells):")
        shown = 0
        for ln in rep.lines:
            if shown >= max_lines:
                break
            cells = "  ".join(f"S{i}={c.text}" for i, c in sorted(ln.cells.items()))
            if cells or ln.label:
                parts.append(f"    L{ln.id} | {ln.label[:60]} | {cells}")
                shown += 1
    return "\n".join(parts)


def decide(grids: list[PageGrid], groups: list[dict], images: list[bytes]) -> dict:
    from openai_backend import decide as openai_decide
    return openai_decide(grids, groups, images)


def rule_for(line, rules: list[dict], position: int, prev_rule: dict | None) -> dict | None:
    """First matching rule; rules tied to the line above ("after") are tried first."""
    prev_id = (prev_rule or {}).get("id")
    for r in sorted(rules, key=lambda r: r.get("after") is None):
        if r.get("after") is not None and r["after"] != prev_id:
            continue
        rx = r.get("label_regex")
        try:
            hit = rx is None or re.search(rx, line.label or "")
        except re.error:
            hit = False
        if not hit or r.get("position") not in (None, position):
            continue
        return r
    return None


def walk(g: PageGrid, rules: list[dict], kinds: list[dict] | None = None):
    """Yield (line, rule, next_rule) for the amount lines of a page, in order. rule is None if nothing matched.
    Lines of kind 'ignore' are not schedule rows and are skipped."""
    seq, position, prev = [], 0, None
    for ln in g.lines:
        if ln.cells and kinds and kind_of(ln.label, kinds) == "ignore":
            continue
        if not ln.cells:
            if ln.label:
                position, prev = 0, None
            continue
        position += 1
        r = rule_for(ln, rules, position, prev)
        if r is not None and r.get("after") is None:
            prev = r
        seq.append((ln, r))
    for i, (ln, r) in enumerate(seq):
        yield ln, r, seq[i + 1][1] if i + 1 < len(seq) else None


def over_ignored(out: dict, grids: list[PageGrid], limit: int = 8) -> list[str]:
    """'ignore' kinds are for page headers and page numbers; a line with three or more amounts is a schedule row."""
    roles = {p.get("page"): p.get("role") for p in out.get("pages", [])}
    hits = [f"page {g.page} line '{ln.label}' ({len(ln.cells)} amounts)" for g in grids
            if roles.get(g.page) == "table" for ln in g.lines
            if len(ln.cells) >= 3 and kind_of(ln.label, out.get("kinds", [])) == "ignore"]
    if not hits:
        return []
    return [f"{len(hits)} lines with several amounts fall under a kind 'ignore' regex, but they look like schedule "
            "rows; make the 'ignore' regexes match only page headers / page numbers, e.g. these must not be "
            "ignored:\n    " + "\n    ".join(hits[:limit])]


def coverage(out: dict, grids: list[PageGrid], groups: list[dict], limit: int = 15) -> list[str]:
    """Amount lines on table pages that no rule matches: the rules are incomplete."""
    roles = {p.get("page"): p.get("role") for p in out.get("pages", [])}
    rules = {g.get("group"): g.get("rules", []) for g in out.get("groups", [])}
    group_of = {pg: gi for gi, gr in enumerate(groups) for pg in gr["pages"]}
    misses = []
    for g in grids:
        if roles.get(g.page) != "table" or g.page not in group_of:
            continue
        for ln, r, _ in walk(g, rules.get(group_of[g.page], []), out.get("kinds", [])):
            if r is None:
                cells = "  ".join(f"S{i}={c.text}" for i, c in sorted(ln.cells.items()))
                misses.append(f"page {g.page} group {group_of[g.page]} line '{ln.label}' | {cells}")
    if not misses:
        return []
    return [f"{len(misses)} amount lines match no rule, e.g.:\n    " + "\n    ".join(misses[:limit])
            + "\n  Make the rules cover them (label regexes are searched, not anchored, and labels can carry "
              "extra words), or say role 'other' for pages that are not part of the schedule."]


def validate(out: dict, groups: list[dict]) -> list[str]:
    problems = []
    for g in out.get("groups", []):
        for k, rule in enumerate(g.get("rules", [])):
            seen: dict[str, list[str]] = {}
            for s, f in rule.get("strips", {}).items():
                if f not in FIELDS:
                    problems.append(f"group {g.get('group')} rule {k}: '{f}' is not a field")
                elif f != "ignore":
                    seen.setdefault(f, []).append(s)
            for f, ss in seen.items():
                if len(ss) > 1:
                    problems.append(f"group {g.get('group')} rule {k}: {', '.join(ss)} are all mapped to '{f}'; "
                                    f"a field may appear once per rule, the others are 'ignore' or another field")
            if rule.get("label_regex"):
                try:
                    re.compile(rule["label_regex"])
                except re.error as e:
                    problems.append(f"group {g.get('group')} rule {k}: bad regex: {e}")
    return problems


def load_env(path: Path) -> None:
    if path.exists():
        for line in path.read_text().splitlines():
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def kind_of(label: str, kinds: list[dict]) -> str:
    for k in kinds:
        try:
            if k.get("label_regex") and re.search(k["label_regex"], label):
                return k["kind"] if k["kind"] in KINDS else "asset"
        except re.error:
            continue
    return "asset"
