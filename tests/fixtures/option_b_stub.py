"""Loopback stub that serves a RECORDED (hand-written) option-B plan.

Plans reference evidence by placeholder instead of document-specific IDs:
"@cell:Sheet!B16", "@row:Sheet:19", "@region:Sheet". The stub resolves each
placeholder against the observations in the actual request payload, so the
plan can only name evidence the pipeline really sent.
"""
from __future__ import annotations

from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from threading import Thread

from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

PLANS = Path(__file__).resolve().parent / "option_b_plans"


def load_plan(name: str) -> dict:
    return json.loads((PLANS / f"{name}.json").read_text())


def resolve(template: object, observations: list[dict]) -> object:
    if isinstance(template, dict):
        return {key: resolve(value, observations) for key, value in template.items()}
    if isinstance(template, list):
        return [resolve(value, observations) for value in template]
    if not isinstance(template, str) or not template.startswith("@"):
        return template
    kind, _, reference = template[1:].partition(":")
    if kind == "cell":
        sheet, coordinate = reference.rsplit("!", 1); letters, row = coordinate_from_string(coordinate)
        hits = [o["id"] for o in observations if o["kind"] in {"cell", "empty_cell"} and o.get("sheet") == sheet and o.get("row") == row and o.get("column") == column_index_from_string(letters)]
    elif kind == "row":
        sheet, row = reference.rsplit(":", 1)
        hits = [o["id"] for o in observations if o["kind"] == "logical_row" and o.get("sheet") == sheet and o.get("row") == int(row)]
    elif kind == "region":
        hits = [o["id"] for o in observations if o["kind"] == "table_region" and o.get("sheet") == reference]
    else:
        raise ValueError(template)
    # Evidence the request does not contain (e.g. another sheet's chunk) stays an unknown ID, as a model would send it.
    return hits[0] if len(hits) == 1 else template


class RecordedPlanHandler(BaseHTTPRequestHandler):
    plan: dict = {"plans": []}
    requests: list[dict] = []

    def do_POST(self):
        request = json.loads(self.rfile.read(int(self.headers["Content-Length"]))); type(self).requests.append(request)
        payload = json.loads(request["input"][1]["content"][0]["text"])
        sheets = {o.get("sheet") for o in payload["observations"] if o["kind"] == "table_region"}
        ids = {o["id"] for o in payload["observations"]}
        plan = {"plans": [p for p in type(self).plan["plans"] if p["region_observation_id"] in ids or p["region_observation_id"].split(":", 1)[1] in sheets]}
        text = json.dumps(resolve(plan, payload["observations"]))
        body = json.dumps({"status": "completed", "output": [{"type": "message", "content": [{"type": "output_text", "text": text}]}]}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, format, *args):
        return


@contextmanager
def recorded_plan_server(plan: dict):
    RecordedPlanHandler.plan = plan; RecordedPlanHandler.requests = []
    server = ThreadingHTTPServer(("127.0.0.1", 0), RecordedPlanHandler); thread = Thread(target=server.serve_forever, daemon=True); thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/v1/responses"
    finally:
        server.shutdown(); thread.join()
