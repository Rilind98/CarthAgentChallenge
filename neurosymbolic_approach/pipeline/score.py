"""Development scorer: our key-shaped output against the answer key. +1 correct, 0 blank, -3 wrong.

Rows are matched by where they are printed (same sheet row, or same page and overlapping line box), so a value
counts only if it sits where the key says it is. Our rows that match nothing but emit values count as wrong.
"""

from __future__ import annotations

import re


def _overlap_ratio(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    area = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    return ix * iy / area


def _kind(k):
    return "subtotal" if k.endswith("_subtotal") else k


def score(doc: dict, key: dict) -> dict:
    ours = doc["rows"]
    taken, details = set(), []
    n = {"correct": 0, "wrong": 0, "blank": 0, "missed": 0, "extra": 0, "spurious": 0}
    kind_ok = kind_n = id_ok = id_n = 0
    for kr in key["rows"]:
        kp = kr["pointer"]
        if "sheet" in kp:
            row = int(re.sub(r"[A-Z]+", "", kp["cell"]))
            cands = [o for o in ours if o["pointer"].get("sheet") == kp["sheet"] and o["pointer"].get("row") == row]
        else:
            cands = sorted(((_overlap_ratio(o["pointer"]["bbox"], kp["bbox"]), o) for o in ours
                            if o["pointer"].get("page") == kp["page"] and o["pointer"].get("bbox")), key=lambda x: -x[0])
            cands = [o for ratio, o in cands if ratio >= 0.5][:1]
        truth = {**kr.get("fields", {}), **(kr.get("net_fields") or {})}
        if not cands:
            n["missed"] += len(truth)
            details.append({"key_row": kr.get("asset_no") or kr.get("class") or kr.get("position"), "kind": kr["kind"],
                            "pointer": kp, "outcome": "row not found", "values": len(truth)})
            continue
        o = cands[0]
        taken.add(o["id"])
        kind_n += 1
        kind_ok += _kind(kr["kind"]) == o["kind"] or (kr["kind"] == "asset" and o["kind"] == "asset")
        if kr["kind"] == "asset" and kr.get("asset_no"):
            id_n += 1
            id_ok += str(o.get("asset_no")) == str(kr["asset_no"]) and str(o.get("sub_no") or "") in (str(kr.get("sub_no") or ""), "")
        for f, v in truth.items():
            got = o["fields"].get(f, "missing")
            if got == "missing" or got is None:
                n["blank"] += 1
                if got == "missing":
                    details.append({"row": o["id"], "field": f, "key": v, "ours": "field not output", "printed": kr.get("printed", {}).get(f)})
            elif got == v:
                n["correct"] += 1
            else:
                n["wrong"] += 1
                details.append({"row": o["id"], "field": f, "key": v, "ours": got, "printed": kr.get("printed", {}).get(f),
                                "our_printed": o["printed"].get(f)})
        for f, v in o["fields"].items():
            if f not in truth and v is not None:
                n["extra"] += 1
                details.append({"row": o["id"], "field": f, "key": "not in key", "ours": v})
    for o in ours:
        if o["id"] not in taken:
            k = sum(1 for v in o["fields"].values() if v is not None)
            n["spurious"] += k
            if k:
                details.append({"row": o["id"], "outcome": "row not in key", "label": o["label"], "values": k})
    total_values = sum(len({**kr.get("fields", {}), **(kr.get("net_fields") or {})}) for kr in key["rows"])
    return {"values_in_key": total_values, **n,
            "score": n["correct"] - 3 * (n["wrong"] + n["extra"] + n["spurious"]),
            "kind_accuracy": round(kind_ok / kind_n, 3) if kind_n else None,
            "asset_id_accuracy": round(id_ok / id_n, 3) if id_n else None,
            "details": details[:60]}
