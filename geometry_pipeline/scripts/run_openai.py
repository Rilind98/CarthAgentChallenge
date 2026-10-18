"""The agent/ pipeline with OpenAI instead of Claude Code in both model steps (A/B against agent/run.py):

  layout  - the original text-in / text-out call (page images + grid description -> JSON), retried on problems
  OCR     - GPT proofreads tesseract's words of each scanned table page in a function-calling loop over the same
            fixed tools Claude gets (agent/ocr_tools.py: view, check, fix, merge, add); view's image is shown to it

Everything else is agent/ unchanged. Proofread pages are cached under ocr_cache/review_gpt/ (never Claude's cache).

Usage: .venv/bin/python scripts/run_openai.py INPUT_DIR OUTPUT_DIR [--only SUBSTR] [--jobs 6]
       OPENAI_MODEL (default gpt-6.1-sol), MAX_CALLS (tool calls per page, default 150)
"""
import base64
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "agent"))

from openai import OpenAI  # noqa: E402

import layout  # noqa: E402
import ocr  # noqa: E402
import run  # noqa: E402

MODEL = os.environ.get("OPENAI_MODEL", "gpt-6.1-sol")
MAX_CALLS = int(os.environ.get("MAX_CALLS", "150"))
ocr.CACHE = ROOT / "ocr_cache" / "review_gpt"


def decide(grids, groups, images):
    """layout.decide as it was before the Claude swap: one chat call, retried with the problems code finds."""
    text = layout.PROMPT.format(fields="\n".join(f"  {k}: {v}" for k, v in layout.FIELDS.items()),
                                doc=layout.describe(grids, groups))
    content = [{"type": "text", "text": text}] + [
        {"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.b64encode(i).decode()}}
        for i in images[:3]]
    client, messages, usage = OpenAI(), [{"role": "user", "content": content}], {"in": 0, "out": 0, "cost_usd": 0.0}
    for _ in range(3):
        r = client.chat.completions.create(model=MODEL, messages=messages, response_format={"type": "json_object"})
        usage["in"] += r.usage.prompt_tokens
        usage["out"] += r.usage.completion_tokens
        out = json.loads(r.choices[0].message.content)
        problems = (layout.validate(out, groups) or layout.over_ignored(out, grids)
                    or layout.coverage(out, grids, groups))
        if not problems:
            break
        messages += [{"role": "assistant", "content": r.choices[0].message.content},
                     {"role": "user", "content": "Fix these problems and answer with the full JSON again:\n- "
                                                 + "\n- ".join(problems)}]
    out.update(_problems=problems, _groups=groups, _usage=usage)
    return out


def fn(name, desc, props, required=()):
    return {"type": "function", "function": {"name": name, "description": desc, "parameters": {
        "type": "object", "properties": props, "required": list(required)}}}


INT = {"type": "integer"}
TOOLS = [
    fn("view", "Image of a page region (pixels; no box = whole page) and its words as `id [box] text`. "
               "Use regions of about 1200 x 600 px; the full page is too small to read digits.",
       {"x0": INT, "y0": INT, "x1": INT, "y1": INT, "boxes": {"type": "boolean", "description": "draw word boxes"}}),
    fn("check", "The row arithmetic, computed in code on the current words: rows that do not add up, their cells, "
                "and per cell the value that alone would make the row add up.", {}),
    fn("fix", "Set a word's text; 'A B' splits it in two; '' deletes it.",
       {"id": INT, "text": {"type": "string"}}, ["id", "text"]),
    fn("merge", "Join pieces of one printed number (text joined left to right unless given).",
       {"ids": {"type": "array", "items": INT}, "text": {"type": "string"}}, ["ids"]),
    fn("add", "Add a missed token with the pixel box of its ink.",
       {"x0": INT, "y0": INT, "x1": INT, "y1": INT, "text": {"type": "string"}}, ["x0", "y0", "x1", "y1", "text"]),
]

# the Claude prompt with its tool section replaced: same task, same rules
PROMPT = ocr.PROMPT.split("You can run only")[0] + "Your tools: view, check, fix, merge, add.\n\n" \
    + "1." + ocr.PROMPT.split("\n1.", 1)[1]


def call_tool(d: Path, name: str, a: dict) -> str:
    if name == "view":
        args = ["view"] + ([str(a[k]) for k in ("x0", "y0", "x1", "y1")] if all(k in a for k in ("x0", "y0", "x1", "y1")) else [])
        args += ["--boxes"] if a.get("boxes") else []
    elif name == "check":
        args = ["check"]
    elif name == "fix":
        args = ["fix", str(a["id"]), a["text"]]
    elif name == "merge":
        args = ["merge", *map(str, a["ids"])] + (["--text", a["text"]] if a.get("text") else [])
    else:
        args = ["add", *(str(a[k]) for k in ("x0", "y0", "x1", "y1")), a["text"]]
    r = subprocess.run([sys.executable, str(ocr.TOOLS), str(d), *args], capture_output=True, text=True)
    return (r.stdout + r.stderr).strip()


def review_page(scan: dict, lay: dict):
    """ocr.review_page with GPT driving the tools."""
    if os.environ.get("OCR_AGENT", "1") == "0":
        return ocr.to_words(scan["tesseract"]), {"cost_usd": 0.0}
    import tempfile
    d = Path(tempfile.mkdtemp(prefix="gpt_review_"))
    (d / "page.png").write_bytes(scan["png"])
    (d / "words.json").write_text(json.dumps(scan["tesseract"]))
    (d / "layout.json").write_text(json.dumps({**lay, "dpi": ocr.DPI}))
    w, h = scan["size"]
    client, usage = OpenAI(), {"in": 0, "out": 0, "cost_usd": 0.0, "calls": 0}
    messages = [{"role": "user", "content": PROMPT.format(w=w, h=h, dpi=ocr.DPI)}]
    answer = ""
    with ocr.LIMIT:
        while usage["calls"] < MAX_CALLS:
            r = client.chat.completions.create(model=MODEL, messages=messages, tools=TOOLS)
            usage["in"] += r.usage.prompt_tokens
            usage["out"] += r.usage.completion_tokens
            m = r.choices[0].message
            messages.append(m.model_dump(exclude_none=True))
            if not m.tool_calls:
                answer = m.content or ""
                break
            images = []
            for tc in m.tool_calls:
                usage["calls"] += 1
                try:
                    out = call_tool(d, tc.function.name, json.loads(tc.function.arguments or "{}"))
                except (KeyError, json.JSONDecodeError) as e:
                    out = f"bad arguments: {e}"
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": out})
                if tc.function.name == "view" and out.startswith("image: "):
                    images.append(d / out.split()[1])
            if images:  # tool results carry text only: show the crops in a user turn
                messages.append({"role": "user", "content": [{"type": "text", "text": "The images of your view calls:"}] + [
                    {"type": "image_url", "image_url": {"url": "data:image/png;base64,"
                                                        + base64.b64encode(p.read_bytes()).decode()}} for p in images]})
    words = json.loads((d / "words.json").read_text())
    edits = json.loads((d / "edits.json").read_text()) if (d / "edits.json").exists() else []
    cache = scan["cache"]
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps({"k": scan["k"], "w": scan["w"], "h": scan["h"], "dpi": ocr.DPI, "answer": answer,
                                 "usage": usage, "edits": edits, "tesseract": scan["tesseract"], "words": words},
                                ensure_ascii=False))
    return ocr.to_words(words), usage


run.decide = decide
run.review_page = review_page

if __name__ == "__main__":
    run.main()
