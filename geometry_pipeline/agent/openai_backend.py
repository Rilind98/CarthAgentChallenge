"""OpenAI layout and OCR harness from the GPT comparison, using the current scan interface.

Direct OpenAI uses Responses; the previously used OpenRouter endpoint uses Chat Completions.
All OCR review passes share this backend and the same fixed, logged correction tools.
"""
import base64
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from openai import OpenAI


def image_url(data: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(data).decode()


class Conversation:
    def __init__(self):
        self.client = OpenAI()
        self.chat = "openrouter.ai" in str(self.client.base_url)
        self.model = os.environ.get("OPENAI_MODEL", "gpt-6.1-sol")
        if self.chat and "/" not in self.model:
            self.model = "openai/" + self.model
        self.messages = []
        self.usage = {"in": 0, "out": 0, "cost_usd": 0.0, "calls": 0}

    def user(self, text: str, images=()):
        if self.chat:
            content = [{"type": "text", "text": text}]
            content += [{"type": "image_url", "image_url": {"url": image_url(i)}} for i in images]
        else:
            content = [{"type": "input_text", "text": text}]
            content += [{"type": "input_image", "image_url": image_url(i)} for i in images]
        self.messages.append({"role": "user", "content": content})

    def request(self, tools=None, json_output=False):
        if self.chat:
            kwargs = {"tools": tools} if tools is not None else {}
            if json_output:
                kwargs["response_format"] = {"type": "json_object"}
            r = self.client.chat.completions.create(model=self.model, messages=self.messages, **kwargs)
            self.usage["in"] += r.usage.prompt_tokens
            self.usage["out"] += r.usage.completion_tokens
            message = r.choices[0].message
            self.messages.append(message.model_dump(exclude_none=True))
            return message.content or "", [(t.id, t.function.name, t.function.arguments)
                                            for t in message.tool_calls or []]
        kwargs = {"store": False, "include": ["reasoning.encrypted_content"]}
        if tools is not None:
            kwargs["tools"] = [{"type": "function", **t["function"], "strict": False} for t in tools]
        if json_output:
            kwargs["text"] = {"format": {"type": "json_object"}}
        r = self.client.responses.create(model=self.model, input=self.messages, **kwargs)
        self.usage["in"] += r.usage.input_tokens
        self.usage["out"] += r.usage.output_tokens
        self.messages.extend(item.model_dump(exclude_none=True) for item in r.output)
        return r.output_text, [(item.call_id, item.name, item.arguments)
                               for item in r.output if item.type == "function_call"]

    def result(self, call_id: str, output: str):
        if self.chat:
            self.messages.append({"role": "tool", "tool_call_id": call_id, "content": output})
        else:
            self.messages.append({"type": "function_call_output", "call_id": call_id, "output": output})


def decide(grids, groups, images):
    import layout
    text = layout.PROMPT.format(fields="\n".join(f"  {k}: {v}" for k, v in layout.FIELDS.items()),
                                doc=layout.describe(grids, groups))
    if not images:
        text = "This document is an Excel workbook: each page is one sheet, each line one row.\n\n" + text
    session = Conversation()
    session.user(text, images[:3])
    for _ in range(3):
        answer, _ = session.request(json_output=True)
        try:
            out = json.loads(answer)
            if not isinstance(out, dict):
                raise ValueError("expected a JSON object")
        except ValueError as e:
            out, problems = {}, [f"answer is not a valid JSON object: {e}"]
        else:
            problems = (layout.validate(out, groups) or layout.over_ignored(out, grids)
                        or layout.coverage(out, grids, groups))
        if not problems:
            break
        session.user("Fix these problems and answer with the full JSON again:\n- " + "\n- ".join(problems))
    out.update(_problems=problems, _groups=groups, _usage=session.usage)
    return out


def fn(name, desc, props, required=()):
    return {"type": "function", "function": {"name": name, "description": desc, "parameters": {
        "type": "object", "properties": props, "required": list(required), "additionalProperties": False}}}


INT = {"type": "integer"}
TOOLS = [
    fn("view", "Image of a page region (pixels; no box = whole page) and its words as id [box] text. "
       "Use regions of about 1200 x 600 px; the full page is too small to read digits.",
       {"x0": INT, "y0": INT, "x1": INT, "y1": INT, "boxes": {"type": "boolean"}}),
    fn("check", "Row arithmetic computed in code: rows that do not add up, their cells, and candidate values. "
       "Candidates are hints only; always confirm against the printed image.", {}),
    fn("fix", "Set a word's text; 'A B' splits it in two; '' deletes it.",
       {"id": INT, "text": {"type": "string"}}, ["id", "text"]),
    fn("merge", "Join pieces of one printed number, left to right unless text is given.",
       {"ids": {"type": "array", "items": INT}, "text": {"type": "string"}}, ["ids"]),
    fn("add", "Add a missed token with the pixel box of its ink.",
       {"x0": INT, "y0": INT, "x1": INT, "y1": INT, "text": {"type": "string"}},
       ["x0", "y0", "x1", "y1", "text"]),
]


def call_tool(directory: Path, name: str, args: dict) -> str:
    import ocr
    if name == "view":
        coords = ("x0", "y0", "x1", "y1")
        cmd = ["view"] + ([str(args[k]) for k in coords] if all(k in args for k in coords) else [])
        cmd += ["--boxes"] if args.get("boxes") else []
    elif name == "check":
        cmd = ["check"]
    elif name == "fix":
        cmd = ["fix", str(args["id"]), args["text"]]
    elif name == "merge":
        cmd = ["merge", *map(str, args["ids"])] + (["--text", args["text"]] if args.get("text") else [])
    elif name == "add":
        cmd = ["add", *(str(args[k]) for k in ("x0", "y0", "x1", "y1")), args["text"]]
    else:
        return f"Unknown tool: {name}"
    result = subprocess.run([sys.executable, str(ocr.TOOLS), str(directory), *cmd],
                            capture_output=True, text=True)
    return (result.stdout + result.stderr).strip()


def ask(scan: dict, layout: dict, prompt: str):
    import ocr
    # Retain the established task and image-only correction rules; translate CLI tools into functions.
    if "\n1." in prompt:
        prompt = prompt.split("You can run only", 1)[0] + "Your tools: view, check, fix, merge, add.\n\n1." \
            + prompt.split("\n1.", 1)[1]
    else:
        prompt = prompt.split("You can run only", 1)[0] + "Use view for close-up images and fix to correct words."
    prompt = prompt.replace("Run check", "Call check")
    session = Conversation()
    max_calls = max(1, int(os.environ.get("MAX_CALLS", "300")))
    # The second pass only needs to view and correct the cells selected by column-sum checks.
    tools = TOOLS if "\n1." in prompt else [t for t in TOOLS if t["function"]["name"] in ("view", "fix")]
    with tempfile.TemporaryDirectory(prefix="gpt_review_") as tmp:
        directory = Path(tmp)
        (directory / "page.png").write_bytes(scan["png"])
        (directory / "words.json").write_text(json.dumps(scan["words"]))
        (directory / "layout.json").write_text(json.dumps({**layout, "dpi": ocr.DPI}))
        session.user(prompt)
        answer = ""
        with ocr.LIMIT:
            while session.usage["calls"] < max_calls:
                answer, calls = session.request(tools=tools)
                if not calls:
                    break
                images = []
                for call_id, name, arguments in calls:
                    if session.usage["calls"] >= max_calls:
                        output = "Tool-call budget exhausted."
                    else:
                        session.usage["calls"] += 1
                        try:
                            output = call_tool(directory, name, json.loads(arguments or "{}"))
                        except (KeyError, ValueError, TypeError) as e:
                            output = f"Bad arguments: {e}"
                    session.result(call_id, output)
                    if name == "view" and output.startswith("image: "):
                        images.append((directory / output.split()[1]).read_bytes())
                if images:
                    session.user("The images of your view calls:", images)
            else:
                raise RuntimeError(f"OCR reached MAX_CALLS={max_calls}; increase it to finish this page")
        words = json.loads((directory / "words.json").read_text())
        edits_path = directory / "edits.json"
        edits = json.loads(edits_path.read_text()) if edits_path.exists() else []
    return answer, words, edits, session.usage
