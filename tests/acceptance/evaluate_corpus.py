"""Independent development-only corpus evaluator; never imported by runtime.

The official scorer is unavailable. This checks signed printed values, fields,
row geometry/cells and available identities against the supplied answer keys.
It reports a local proxy, never leaderboard compatibility.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import re

# Public semantic aliases, independent of any document family or value.
FIELD_ALIASES = {"umb": "umbuchung", "total_afa_gj": "afa_gj"}


def printed_decimal(raw: str, scale: Decimal) -> Decimal | None:
    value = str(raw).strip().replace("\u00a0", "").replace(" ", "").replace("−", "-")
    if value == "":
        return Decimal(0)
    negative = value.startswith("(") and value.endswith(")") or value.endswith("-") or value.startswith("-")
    value = value.strip("()-+")
    if not re.fullmatch(r"\d+(?:[.,]\d+)*", value):
        return None
    if "," in value and "." in value:
        decimal_mark = "," if value.rfind(",") > value.rfind(".") else "."
        value = value.replace("." if decimal_mark == "," else ",", "").replace(decimal_mark, ".")
    elif "," in value:
        value = value.replace(",", ".")
    elif value.count(".") > 1 or "." in value and len(value.rsplit(".", 1)[1]) == 3:
        value = value.replace(".", "")
    try:
        return Decimal(value) * (-1 if negative else 1) * scale
    except InvalidOperation:
        return None


def expected_values(key):
    scale = Decimal(1000 if str(key["meta"].get("unit", "EUR")).upper() in {"TEUR", "T€"} else 1)
    result = []
    for index, row in enumerate(key["rows"]):
        fields = {**row.get("fields", {}), **row.get("net_fields", {})}
        for field, cents in fields.items():
            pointer = row.get("pointer", {})
            # The key sometimes carries an accounting zero for a nonprinted
            # subtotal field. With no physical XLSX cell it is not a source slot.
            if "sheet" in pointer and field not in pointer.get("cells", {}):
                continue
            printed = row.get("printed", {}).get(field)
            value = printed_decimal(printed, scale) if printed is not None else Decimal(cents) / 100
            result.append({"row": index, "field": field, "amount": value, "pointer": pointer,
                           "asset": row.get("asset_no"), "subnumber": row.get("sub_no"), "class": row.get("class"), "name": row.get("name"),
                           "year": str(row.get("year", "")), "kind": row.get("kind"),
                           "printed": printed, "flag": row.get("flags", {}).get(field)})
    return result


def location_matches(fact, gold, *, tolerance=True):
    p, g = fact["pointer"], gold["pointer"]
    if "sheet" in g:
        return p.get("sheet") == g["sheet"] and p.get("cell") == g.get("cells", {}).get(gold["field"])
    if p.get("page_index", -2) + 1 != g.get("page"):
        return False
    box = p.get("bbox")
    row = g.get("bbox")
    if not box or not row:
        return False
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    dx, dy = (2, 3) if tolerance else (0, 0)
    return row[0] - dx <= cx <= row[2] + dx and row[1] - dy <= cy <= row[3] + dy


def matching_candidates(fact, expected):
    semantic_field = FIELD_ALIASES.get(fact["field"], fact["field"])
    candidates = [(i, g) for i, g in enumerate(expected) if g["field"] == semantic_field and location_matches(fact, g)]
    if len(candidates) > 1:
        # Expanded row tolerances can overlap on tightly spaced scans. Prefer
        # unique containment in the original gold row, without using the value
        # or asset identity to choose a row. Truly ambiguous boxes stay ambiguous.
        exact = [(i, g) for i, g in candidates if location_matches(fact, g, tolerance=False)]
        if len(exact) == 1:
            return exact
    return candidates


def evaluate(corpus: Path, output: Path):
    keys = {}
    for key_path in sorted(corpus.glob("f*/*.key.json")):
        key = json.loads(key_path.read_text())
        stem = key_path.name.removesuffix(".key.json")
        candidates = [key_path.with_name(stem + ext) for ext in (".pdf", ".xlsx")]
        source = next((p for p in candidates if p.exists()), None)
        if source:
            digest = hashlib.sha256(source.read_bytes()).hexdigest()
            keys[digest] = (key, source)
    payloads = defaultdict(list)
    for path in output.rglob("evidence.v1.json"):
        if not (path.parent / ".complete").exists():
            continue
        value = json.loads(path.read_text())
        payloads[value["document"]["source_hash"]].append(value)
    totals = Counter()
    formats = defaultdict(Counter)
    families = defaultdict(Counter)
    documents = []
    for digest, (key, source) in keys.items():
        m = Counter(documents=1)
        expected = expected_values(key)
        m["source_slots"] = len(expected)
        records = payloads.get(digest, [])
        mismatches = []
        if len(records) != 1:
            m["missing_output" if not records else "multiple_output_versions"] = 1
            m["unaccepted_slots"] = len(expected)
        else:
            record = records[0]
            m["published_outputs"] = 1
            m["status_" + record["document"]["status"]] = 1
            matched = set()
            for fact in record["facts"]:
                m["accepted"] += 1
                candidates = matching_candidates(fact, expected)
                correct = False
                reason = "field_or_pointer"
                if len(candidates) == 1:
                    index, gold = candidates[0]
                    source_value = Decimal(fact["amount"]["normalized_base"])
                    actual_identity = str(fact["asset_identity"].get("asset_or_class", ""))
                    actual_subnumber = str(fact["asset_identity"].get("asset_sub_id", ""))
                    actual_period = str(fact["asset_identity"].get("period", ""))
                    if gold["amount"] is None:
                        reason = "unparseable_printed_gold"
                        m["unevaluable_accepted"] += 1
                        continue
                    if source_value != gold["amount"]:
                        reason = "signed_source_value"
                    elif gold["asset"] and actual_identity != str(gold["asset"]):
                        reason = "asset_identity"
                    elif gold["subnumber"] not in (None, "") and actual_subnumber != str(gold["subnumber"]):
                        reason = "asset_subnumber"
                    elif gold["year"] and gold["year"] not in actual_period:
                        reason = "period"
                    elif index in matched:
                        reason = "duplicate_acceptance"
                    else:
                        correct = True
                        matched.add(index)
                        m["correct_zero" if source_value == 0 else "correct_nonzero"] += 1
                if correct:
                    m["correct"] += 1
                else:
                    m["wrong"] += 1
                    m["wrong_" + reason] += 1
                    if len(mismatches) < 12:
                        mismatches.append({"id": fact["id"], "field": fact["field"], "pointer": fact["pointer"], "reason": reason})
            m["unaccepted_slots"] = len(expected) - len(matched)
            m["reported_abstentions"] = len(record["abstentions"])
        m["proxy_score"] = m["correct"] - 3 * m["wrong"]
        meta = key["meta"]
        totals.update(m)
        formats[meta["format"]].update(m)
        families[meta["family"]].update(m)
        documents.append({"source": source.name, "format": meta["format"], "family": meta["family"], "metrics": dict(m), "mismatch_examples": mismatches})
    def add_rates(counter):
        result = dict(counter)
        result["accepted_precision"] = counter["correct"] / (counter["correct"] + counter["wrong"]) if counter["correct"] + counter["wrong"] else None
        result["correct_coverage"] = counter["correct"] / counter["source_slots"] if counter["source_slots"] else None
        return result
    return {"scope": "Synthetic development proxy, not the official scorer. PDF gold supplies row boxes, not exact value boxes; independent pointer tests are also required.",
            "pointer_policy": "Exact XLSX cells; PDF amount centers within gold row boxes with 2pt horizontal/3pt vertical tolerance. Unique containment in an unexpanded row resolves overlapping tolerances; value and identity never select the row.",
            "sign_policy": "Compare signed printed values and source scale, not blind absolute-value core-key magnitudes.",
            "field_aliases": FIELD_ALIASES,
            "totals": add_rates(totals), "by_format": {k: add_rates(v) for k,v in formats.items()},
            "by_family": {k: add_rates(v) for k,v in families.items()}, "documents": documents}


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("corpus", type=Path)
    p.add_argument("outputs", type=Path)
    p.add_argument("--report", type=Path, required=True)
    args = p.parse_args()
    report = evaluate(args.corpus, args.outputs)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"totals": report["totals"], "by_format": report["by_format"]}, indent=2))
