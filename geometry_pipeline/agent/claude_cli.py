"""Ask Claude Code (headless `claude -p`, local OAuth login) and get its answer back.

Each conversation runs in its own temporary directory that holds the files Claude may read (page images, OCR
words). A follow-up resumes the same session, so Claude keeps what it already looked at. `tools` are the built-in
tools it gets; `allow` narrows them to exact calls (e.g. "Bash(./tool *)") and denies everything else.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

MODEL = os.environ.get("CLAUDE_MODEL", "opus")


class Session:
    def __init__(self, files: dict[str, bytes | str] | None = None, tools: str = "Read",
                 allow: list[str] | None = None):
        self.dir = Path(tempfile.mkdtemp(prefix="claude_"))
        for name, data in (files or {}).items():
            (self.dir / name).write_bytes(data if isinstance(data, bytes) else data.encode())
        self.tools = tools
        self.allow = allow  # permission rules, e.g. "Bash(./tool *)": only these calls run, all others are denied
        self.id: str | None = None
        self.usage = {"in": 0, "out": 0, "cost_usd": 0.0}

    def ask(self, prompt: str, timeout: int = 1800) -> str:
        env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")}
        cmd = ["claude", "-p", prompt, "--model", MODEL, "--output-format", "json", "--tools", self.tools]
        if self.allow:
            cmd += ["--allowedTools", ",".join(self.allow), "--permission-mode", "dontAsk"]
        else:
            cmd += ["--allowedTools", self.tools, "--permission-mode", "bypassPermissions"]
        if self.id:
            cmd += ["--resume", self.id]
        r = subprocess.run(cmd, cwd=self.dir, env=env, capture_output=True, text=True, timeout=timeout)
        try:
            out = json.loads(r.stdout)
        except json.JSONDecodeError:
            raise RuntimeError(f"claude failed ({r.returncode}): {(r.stdout + r.stderr)[-500:]}")
        if out.get("is_error"):
            raise RuntimeError(f"claude error: {out.get('result', '')[:500]}")
        self.id = out["session_id"]
        u = out.get("usage", {})
        self.usage["in"] += u.get("input_tokens", 0) + u.get("cache_read_input_tokens", 0) \
            + u.get("cache_creation_input_tokens", 0)
        self.usage["out"] += u.get("output_tokens", 0)
        self.usage["cost_usd"] += out.get("total_cost_usd", 0.0)
        return out["result"]

    def close(self) -> None:
        shutil.rmtree(self.dir, ignore_errors=True)


def parse_json(text: str):
    """The JSON object in an answer (Claude may wrap it in a ```json fence or a sentence)."""
    m = re.search(r"```(?:json)?\s*(\{.*\}|\[.*\])\s*```", text, re.S)
    if m:
        return json.loads(m.group(1))
    start = min((i for i in (text.find("{"), text.find("[")) if i >= 0), default=-1)
    end = max(text.rfind("}"), text.rfind("]"))
    if start < 0 or end < start:
        raise ValueError(f"no JSON in answer: {text[:300]}")
    return json.loads(text[start:end + 1])
