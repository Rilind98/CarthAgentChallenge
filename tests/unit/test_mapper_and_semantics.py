from __future__ import annotations

from dataclasses import replace
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import sqlite3
from threading import Thread
import time

from openpyxl import Workbook
import pytest

from carthagent.mapper import MapperConfig, MapperUnavailableError, MapperValidationError, OpenAICompatibleMapper
from carthagent.models import CoverageDisposition, CoverageItem, SourceObservation
from carthagent.policies import DEFAULT_BOUNDS, DEFAULT_POLICIES
from carthagent.pipeline import RunConfig, _mapper_bounded_chunks, run_input
from carthagent.semantics import interpret


def xobs(cell: str, text: str, kind: str = "cell", metadata=None):
    return SourceObservation(f"d:x:{cell}", "d", "a" * 64, kind, f"Data!{cell}", "xlsx_native" if kind == "cell" else "xlsx_structure", text, metadata=metadata or {})


def table_observations(*, field_header: str = "Custom metric EUR 2025"):
    asset_header = xobs("A1", "Asset ID"); amount_header = xobs("B1", field_header)
    asset = xobs("A2", "A-1"); amount = xobs("B2", "123")
    row = xobs("row:2", "A-1 | 123", "logical_row", {"member_ids": [asset.id, amount.id], "sheet": "Data", "row": 2})
    region = xobs("region", "Data", "table_region", {"member_ids": [row.id], "sheet": "Data"})
    return asset_header, amount_header, asset, amount, row, region


def plan(region, field_header, asset_header, row, *, semantic_field="ahk_open"):
    return {"plans": [{
        "region_observation_id": region.id,
        "field_headers": [{"observation_id": field_header.id, "semantic_field": semantic_field, "parent_header_ids": []}],
        "identity_headers": [{"observation_id": asset_header.id, "identity_role": "asset_id"}],
        "measure_axes": [],
        "movement_axes": [],
        "row_roles": [{"observation_id": row.id, "row_role": "data"}],
        "period_contexts": [{"observation_id": field_header.id, "temporal_role": "current_period"}],
        "unit_contexts": [{"observation_id": field_header.id, "unit": "EUR"}],
    }]}


def test_generic_and_unknown_fields_both_abstain_until_refined():
    observations = table_observations(field_header="Amount EUR 2025")
    asset_header, amount_header, asset, amount, row, region = observations
    coverage = (CoverageItem("d:c", "d", "cell", "Data!B2", CoverageDisposition.UNRESOLVED),)
    _, facts, abstentions, _ = interpret(observations, coverage, DEFAULT_POLICIES)
    assert not facts and abstentions[0].reason.startswith("missing_context:field")
    mystery = xobs("B1", "Mystery EUR 2025")
    unknown = (asset_header, mystery, asset, amount, row, region)
    _, facts, abstentions, _ = interpret(unknown, coverage, DEFAULT_POLICIES)
    assert not facts and abstentions[0].reason.startswith("missing_context:field")


class StubHandler(BaseHTTPRequestHandler):
    response_content = {}
    request_json: dict | None = None

    def do_POST(self):
        length = int(self.headers["Content-Length"]); type(self).request_json = json.loads(self.rfile.read(length))
        content = type(self).response_content
        if callable(content):
            content = content(type(self).request_json)
        body = json.dumps({"status": "completed", "output": [{"type": "message", "content": [{"type": "output_text", "text": json.dumps(content)}]}]}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, format, *args):
        return


@pytest.fixture
def mapper_server():
    server = ThreadingHTTPServer(("127.0.0.1", 0), StubHandler); thread = Thread(target=server.serve_forever, daemon=True); thread.start()
    yield f"http://127.0.0.1:{server.server_port}/v1/responses"
    server.shutdown(); thread.join()


def test_mapper_compact_plan_is_consumed_with_geometry_and_strict_ids(mapper_server):
    observations = table_observations(); asset_header, unknown_header, asset, amount, row, region = observations
    coverage = (CoverageItem("d:c", "d", "cell", "Data!B2", CoverageDisposition.UNRESOLVED),)
    assert not interpret(observations, coverage, DEFAULT_POLICIES)[1]
    StubHandler.response_content = plan(region, unknown_header, asset_header, row)
    mapper = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256))
    mapped = mapper.map("d", observations)
    sent = StubHandler.request_json
    payload = json.loads(sent["input"][1]["content"][0]["text"])
    sent_amount = next(item for item in payload["observations"] if item["id"] == amount.id)
    assert sent_amount["text"] is None and sent_amount["numeric_value"] is True and sent_amount["column"] == 2
    assert sent_amount["row_block_id"] == row.id and sent_amount["region_id"] == region.id
    assert "header_ancestry_ids" not in sent_amount
    assert len(mapped) == 1 and mapped[0].candidates[0].semantic_identity_assignments == ((asset.id, "asset_id"),)
    _, facts, abstentions, _ = interpret(observations, coverage, DEFAULT_POLICIES, mapped)
    assert len(facts) == 1 and not abstentions and facts[0].business.field == "ahk_open" and facts[0].business.asset_id == "A-1"
    assert mapped[0].id in facts[0].dependency_ids
    assert mapper.last_audit and mapper.last_audit["validated_plan"] == StubHandler.response_content
    assert mapper.last_audit["fingerprint"]["model"] == "local-stub" and mapper.last_audit["source_fingerprint"]
    assert sent["model"] == "local-stub" and sent["store"] is False and sent["tools"] == [] and sent["text"]["format"]["strict"] is True


def test_mapper_can_refine_generic_field_but_cannot_override_specific_field(mapper_server):
    observations = table_observations(field_header="Amount EUR 2025"); asset_header, header, _, amount, row, region = observations
    coverage = (CoverageItem("d:c", "d", "cell", amount.locator, CoverageDisposition.UNRESOLVED),)
    assert not interpret(observations, coverage, DEFAULT_POLICIES)[1]
    StubHandler.response_content = plan(region, header, asset_header, row, semantic_field="abgang")
    mapped = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256)).map("d", observations)
    facts = interpret(observations, coverage, DEFAULT_POLICIES, mapped)[1]
    assert len(facts) == 1 and facts[0].business.field == "abgang"

    specific = table_observations(field_header="Additions EUR 2025"); asset_header, header, _, amount, row, region = specific
    StubHandler.response_content = plan(region, header, asset_header, row, semantic_field="abgang")
    mapped = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256)).map("d", specific)
    facts = interpret(specific, coverage, DEFAULT_POLICIES, mapped)[1]
    assert len(facts) == 1 and facts[0].business.field == "zugang"


def test_mapper_typed_measure_and_movement_axes_expand_one_region_plan(mapper_server):
    title = xobs("A1", "Fiscal period 2025 EUR")
    asset_header = xobs("A2", "Asset ID"); measure_header = xobs("B2", "Measure")
    opening_header = xobs("C2", "Opening 2024"); closing_header = xobs("D2", "Closing 2025")
    asset = xobs("A3", "A-1"); measure = xobs("B3", "Unfamiliar basis")
    opening = xobs("C3", "100"); closing = xobs("D3", "110")
    row = xobs("row:3", "A-1 | Unfamiliar basis | 100 | 110", "logical_row", {"member_ids": [asset.id, measure.id, opening.id, closing.id], "sheet": "Data", "row": 3})
    region = xobs("region", "Data", "table_region", {"member_ids": [row.id], "sheet": "Data"})
    observations = (title, asset_header, measure_header, opening_header, closing_header, asset, measure, opening, closing, row, region)
    StubHandler.response_content = {"plans": [{
        "region_observation_id": region.id,
        "field_headers": [],
        "identity_headers": [{"observation_id": asset_header.id, "identity_role": "asset_id"}],
        "measure_axes": [{"observation_id": measure.id, "semantic_measure": "ahk"}],
        "movement_axes": [
            {"observation_id": opening_header.id, "movement_axis": "opening", "parent_header_ids": []},
            {"observation_id": closing_header.id, "movement_axis": "closing", "parent_header_ids": []},
        ],
        "row_roles": [{"observation_id": row.id, "row_role": "data"}],
        "period_contexts": [{"observation_id": closing_header.id, "temporal_role": "current_period"}],
        "unit_contexts": [{"observation_id": title.id, "unit": "EUR"}],
    }]}
    mapper = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256))
    mapped = mapper.map("d", observations)
    coverage = tuple(CoverageItem(f"d:c:{item.id}", "d", "cell", item.locator, CoverageDisposition.UNRESOLVED) for item in (opening, closing))
    facts = interpret(observations, coverage, DEFAULT_POLICIES, mapped)[1]
    assert {(fact.business.field, fact.amount.normalized_value) for fact in facts} == {("ahk_open", 100), ("ahk_close", 110)}
    assert all(fact.business.asset_id == "A-1" and fact.business.period == "2025" for fact in facts)
    assert mapper.last_audit and mapper.last_audit["validated_plan"] == StubHandler.response_content


def test_mapper_row_role_cannot_promote_group_empty_zero(mapper_server):
    asset_header = xobs("A1", "Asset ID"); amount_header = xobs("B1", "Mystery EUR")
    label = xobs("A2", "000101"); empty = xobs("B2", None, "empty_cell", {"confirmed_empty": True, "full_region_present": True})
    row = xobs("row:2", "000101", "logical_row", {"member_ids": [label.id, empty.id], "sheet": "Data", "row": 2})
    region = xobs("region", "Data", "table_region", {"member_ids": [row.id], "sheet": "Data"})
    observations = (asset_header, amount_header, label, empty, row, region)
    StubHandler.response_content = plan(region, amount_header, asset_header, row, semantic_field="zugang")
    mapped = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256)).map("d", observations)
    coverage = (CoverageItem("d:c", "d", "cell", empty.locator, CoverageDisposition.UNRESOLVED),)
    _, facts, _, _ = interpret(observations, coverage, DEFAULT_POLICIES, mapped)
    assert facts == ()


def test_mapper_rejects_foreign_ids_extra_fields_and_insecure_endpoint(mapper_server):
    observations = table_observations(); asset_header, header, _, _, row, region = observations
    mapper = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256))
    foreign = plan(region, header, asset_header, row); foreign["plans"][0]["region_observation_id"] = "other:region"
    StubHandler.response_content = foreign
    with pytest.raises(MapperValidationError): mapper.map("d", observations)
    extra = plan(region, header, asset_header, row); extra["plans"][0]["amount"] = "999"
    StubHandler.response_content = extra
    with pytest.raises(MapperValidationError): mapper.map("d", observations)
    with pytest.raises(MapperValidationError, match="loopback"):
        OpenAICompatibleMapper(MapperConfig(True, "http://api.openai.com/v1/responses", "m", 1, 1)).map("d", observations)


def test_mapper_disabled_makes_zero_calls():
    assert OpenAICompatibleMapper(MapperConfig()).map("d", table_observations()) == ()


def test_mapper_payload_is_linear_and_uses_region_membership_instead_of_repeated_ancestry(mapper_server):
    words = tuple(
        SourceObservation(f"d:w:{index}", "d", "a" * 64, "word", f"page:1:w:{index}", "pdf_native", str(index), (float(index), 10.0, float(index + 1), 20.0), {"page_index": 0, "page_width": 1000, "page_height": 1000})
        for index in range(300)
    )
    row = SourceObservation("d:row", "d", "a" * 64, "logical_row", "page:1:row", "pdf_structure", "row", (0, 10, 300, 20), {"page_index": 0, "structure_role": "visual_row", "member_ids": [word.id for word in words]})
    region = SourceObservation("d:region", "d", "a" * 64, "table_region", "page:1:region", "pdf_structure", "page", metadata={"page_index": 0, "member_ids": [row.id]})
    StubHandler.response_content = {"plans": []}
    mapper = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256))
    assert mapper.map("d", (*words, row, region)) == ()
    payload = json.loads(StubHandler.request_json["input"][1]["content"][0]["text"])
    sent_words = [item for item in payload["observations"] if item["kind"] == "word"]
    assert len(sent_words) == len(words)
    assert all(item["row_block_id"] == row.id and item["region_id"] == region.id for item in sent_words)
    assert all("header_ancestry_ids" not in item for item in payload["observations"])


class ProviderErrorHandler(BaseHTTPRequestHandler):
    responses: list[tuple[int, dict, dict[str, str]]] = []
    calls = 0

    def do_POST(self):
        type(self).calls += 1
        self.rfile.read(int(self.headers["Content-Length"]))
        status, payload, headers = type(self).responses.pop(0)
        body = json.dumps(payload).encode()
        self.send_response(status)
        for name, value in headers.items():
            self.send_header(name, value)
        self.send_header("Content-Type", "application/json"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, format, *args):
        return


@pytest.fixture
def provider_error_server():
    ProviderErrorHandler.calls = 0; ProviderErrorHandler.responses = []
    server = ThreadingHTTPServer(("127.0.0.1", 0), ProviderErrorHandler); thread = Thread(target=server.serve_forever, daemon=True); thread.start()
    yield f"http://127.0.0.1:{server.server_port}/v1/responses"
    server.shutdown(); thread.join()


def test_mapper_http_diagnostics_are_safe_and_nontransient_400_is_not_retried(provider_error_server):
    ProviderErrorHandler.responses = [(400, {"error": {"code": "context_length_exceeded", "type": "invalid_request_error", "param": "input", "message": "SECRET_DOCUMENT_TEXT"}}, {"x-request-id": "req-123"})]
    mapper = OpenAICompatibleMapper(MapperConfig(True, provider_error_server, "local-stub"))
    with pytest.raises(MapperUnavailableError) as caught:
        mapper.map("d", table_observations(), transient_retries=2)
    diagnostic = mapper.last_provider_diagnostic
    assert ProviderErrorHandler.calls == 1
    assert diagnostic == {"status": 400, "message": "mapper request exceeded the provider context window", "code": "context_length_exceeded", "type": "invalid_request_error", "param": "input", "request_id": "req_123"}
    assert "SECRET_DOCUMENT_TEXT" not in str(caught.value)


def test_mapper_retries_only_allowlisted_transient_provider_errors(provider_error_server):
    completed = {"status": "completed", "output": [{"type": "message", "content": [{"type": "output_text", "text": json.dumps({"plans": []})}]}]}
    ProviderErrorHandler.responses = [
        (503, {"error": {"code": "server_error", "message": "internal details"}}, {"Retry-After": "0"}),
        (200, completed, {}),
    ]
    mapper = OpenAICompatibleMapper(MapperConfig(True, provider_error_server, "local-stub"))
    assert mapper.map("d", table_observations(), transient_retries=2) == ()
    assert ProviderErrorHandler.calls == 2

    ProviderErrorHandler.calls = 0
    ProviderErrorHandler.responses = [(429, {"error": {"code": "insufficient_quota", "message": "billing details"}}, {})]
    with pytest.raises(MapperUnavailableError):
        mapper.map("d", table_observations(), transient_retries=2)
    assert ProviderErrorHandler.calls == 1


def test_mapper_retry_after_above_bound_returns_retry_later(provider_error_server):
    ProviderErrorHandler.responses = [(429, {"error": {"code": "slow_down"}}, {"Retry-After": "31"})]
    mapper = OpenAICompatibleMapper(MapperConfig(True, provider_error_server, "local-stub"))
    with pytest.raises(MapperUnavailableError, match="retry_later"):
        mapper.map("d", table_observations(), transient_retries=2)
    assert ProviderErrorHandler.calls == 1


def test_mapper_response_size_limit_rejects_declared_oversize_without_body_leak(mapper_server):
    StubHandler.response_content = {"plans": [], "padding": "secret" * 200}
    mapper = OpenAICompatibleMapper(MapperConfig(True, mapper_server, "local-stub", 2, 256))
    with pytest.raises(MapperUnavailableError, match="mapper_response_size_limit") as caught:
        mapper.map("d", table_observations(), max_response_bytes=64)
    assert "secret" not in str(caught.value)


def test_mapper_hypothesis_dependencies_are_persisted_for_replay(tmp_path, mapper_server):
    source = tmp_path / "mapped.xlsx"
    workbook = Workbook(); sheet = workbook.active; sheet.title = "Data"
    sheet.append(["Fixed asset schedule fiscal 2025 EUR"])
    sheet.append(["Asset ID", "Custom metric EUR"])
    sheet.append(["A-1", 123]); workbook.save(source)

    def response(request):
        payload = json.loads(request["input"][1]["content"][0]["text"])
        observations = payload["observations"]
        by_text = {item.get("text"): item["id"] for item in observations if item.get("text")}
        title = next(item for item in observations if item["kind"] == "cell" and item.get("row") == 1 and item.get("column") == 1)
        row = next(item for item in observations if item["kind"] == "logical_row" and item.get("row") == 3)
        region = next(item for item in observations if item["kind"] == "table_region")
        return {"plans": [{
            "region_observation_id": region["id"],
            "field_headers": [{"observation_id": by_text["Custom metric EUR"], "semantic_field": "ahk_open", "parent_header_ids": []}],
            "identity_headers": [{"observation_id": by_text["Asset ID"], "identity_role": "asset_id"}],
            "measure_axes": [], "movement_axes": [],
            "row_roles": [{"observation_id": row["id"], "row_role": "data"}],
            "period_contexts": [{"observation_id": title["id"], "temporal_role": "current_period"}],
            "unit_contexts": [{"observation_id": title["id"], "unit": "EUR"}],
        }]}

    StubHandler.response_content = response
    config = RunConfig(mapper=MapperConfig(True, mapper_server, "local-stub", 2, 256))
    result = run_input(source, tmp_path / "out", config)
    assert result.status.value in {"complete", "partial"}
    database = sqlite3.connect(tmp_path / "out" / ".carthagent" / "ledger.sqlite3")
    record_ids = {row[0] for row in database.execute("SELECT id FROM records")}
    mapper_ids = {row[0] for row in database.execute("SELECT id FROM records WHERE kind='hypothesis' AND id LIKE '%:mapper:%'")}
    fact_dependencies = {row[0] for row in database.execute("SELECT dependency_id FROM dependencies JOIN records ON records.id=dependencies.record_id WHERE records.kind='fact'")}
    database.close()
    assert mapper_ids and mapper_ids <= fact_dependencies and fact_dependencies <= record_ids


def test_mapper_oversized_region_splits_only_at_complete_rows_and_repeats_context():
    row_count = 1_050
    values = tuple(
        SourceObservation(f"d:v:{index}", "d", "a" * 64, "word", f"page:1:v:{index}", "pdf_native", "X" * 300, (10, float(index), 20, float(index + 1)), {"page_index": 0, "page_width": 1000, "page_height": 2000})
        for index in range(row_count)
    )
    rows = tuple(
        SourceObservation(f"d:r:{index}", "d", "a" * 64, "logical_row", f"page:1:r:{index}", "pdf_structure", "row", (10, float(index), 20, float(index + 1)), {"page_index": 0, "structure_role": "visual_row", "member_ids": [values[index].id]})
        for index in range(row_count)
    )
    region = SourceObservation("d:region:large", "d", "a" * 64, "table_region", "page:1:region", "pdf_structure", "page", metadata={"page_index": 0, "member_ids": [row.id for row in rows]})
    mapper = OpenAICompatibleMapper(MapperConfig())
    chunks = _mapper_bounded_chunks(mapper, "d", (*values, *rows, region), 20)
    assert len(chunks) > 1 and all(mapper.request_fits("d", chunk) for chunk in chunks)
    member_sets = [set(next(item for item in chunk if item.kind == "table_region").metadata["member_ids"]) for chunk in chunks]
    assert set(row.id for row in rows) == set().union(*member_sets)
    assert all(set(row.id for row in rows[:20]) <= members for members in member_sets)
    assert {item.id for item in (*values, *rows, region)} == set().union(*({item.id for item in chunk} for chunk in chunks))
    for chunk, members in zip(chunks, member_sets):
        ids = {item.id for item in chunk}
        assert all(row.id in ids and values[index].id in ids for index, row in enumerate(rows) if row.id in members)


class SlowMapperHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        self.rfile.read(int(self.headers["Content-Length"]))
        time.sleep(2)
        try:
            self.send_response(200); self.end_headers(); self.wfile.write(b"{}")
        except BrokenPipeError:
            pass

    def log_message(self, format, *args):
        return


def test_slow_mapper_budget_preserves_local_facts_and_finishes_with_reserve(tmp_path: Path):
    source = tmp_path / "mixed.xlsx"
    workbook = Workbook(); sheet = workbook.active
    sheet.append(["Fixed asset schedule fiscal 2025 EUR"])
    sheet.append(["Asset ID", "Opening cost EUR", "Mystery EUR"])
    sheet.append(["A-1", 100, 999]); workbook.save(source)
    server = ThreadingHTTPServer(("127.0.0.1", 0), SlowMapperHandler)
    thread = Thread(target=server.serve_forever, daemon=True); thread.start()
    try:
        bounds = replace(DEFAULT_BOUNDS, document_timeout_seconds=3, mapper_phase_timeout_seconds=1, post_mapper_reserve_seconds=2, max_mapper_requests=1)
        config = RunConfig(bounds=bounds, mapper=MapperConfig(True, f"http://127.0.0.1:{server.server_port}/v1/responses", "local-stub", 10, 256))
        result = run_input(source, tmp_path / "out", config)
    finally:
        server.shutdown(); thread.join()
    assert result.status.value == "partial" and result.duration_seconds < bounds.document_timeout_seconds
    assert any("mapper_unavailable" in diagnostic or "mapper_budget_exhausted" in diagnostic for diagnostic in result.diagnostics)
    evidence = json.loads(result.evidence_path.read_text())
    assert any(fact["field"] == "ahk_open" and fact["amount"]["normalized_base"] == "100" for fact in evidence["facts"])


class AnthropicStubHandler(BaseHTTPRequestHandler):
    response_content: object = {}
    stop_reason = "end_turn"
    request_json: dict | None = None
    request_headers: dict[str, str] = {}

    def do_POST(self):
        length = int(self.headers["Content-Length"]); type(self).request_json = json.loads(self.rfile.read(length))
        type(self).request_headers = {name.lower(): value for name, value in self.headers.items()}
        content = [{"type": "thinking", "thinking": "", "signature": "sig"}, {"type": "text", "text": json.dumps(type(self).response_content)}]
        body = json.dumps({"type": "message", "role": "assistant", "content": content, "stop_reason": type(self).stop_reason}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, format, *args):
        return


@pytest.fixture
def anthropic_server():
    AnthropicStubHandler.stop_reason = "end_turn"
    server = ThreadingHTTPServer(("127.0.0.1", 0), AnthropicStubHandler); thread = Thread(target=server.serve_forever, daemon=True); thread.start()
    yield f"http://127.0.0.1:{server.server_port}/v1/messages"
    server.shutdown(); thread.join()


def anthropic_config(endpoint: str, **changes) -> MapperConfig:
    return replace(MapperConfig(True, endpoint, "claude-stub", 2, 256), provider="anthropic", **changes)


def test_anthropic_mapper_sends_messages_api_payload_and_consumes_same_plan(anthropic_server):
    observations = table_observations(); asset_header, header, asset, amount, row, region = observations
    coverage = (CoverageItem("d:c", "d", "cell", "Data!B2", CoverageDisposition.UNRESOLVED),)
    AnthropicStubHandler.response_content = plan(region, header, asset_header, row)
    mapper = OpenAICompatibleMapper(anthropic_config(anthropic_server))
    mapped = mapper.map("d", observations)
    sent = AnthropicStubHandler.request_json
    assert set(sent) == {"model", "max_tokens", "system", "messages", "output_config"}
    assert sent["model"] == "claude-stub" and sent["max_tokens"] == 256 and "thinking" not in sent
    assert sent["output_config"] == {"format": {"type": "json_schema", "schema": OpenAICompatibleMapper.response_schema()}, "effort": "max"}
    assert sent["messages"][0]["role"] == "user" and json.loads(sent["messages"][0]["content"])["document_id"] == "d"
    assert AnthropicStubHandler.request_headers["anthropic-version"] == "2023-06-01"
    assert "x-api-key" not in AnthropicStubHandler.request_headers and "authorization" not in AnthropicStubHandler.request_headers
    facts = interpret(observations, coverage, DEFAULT_POLICIES, mapped)[1]
    assert len(facts) == 1 and facts[0].business.field == "ahk_open" and facts[0].business.asset_id == "A-1"
    assert mapper.last_audit["fingerprint"]["provider"] == "anthropic"


def test_anthropic_mapper_still_rejects_foreign_ids(anthropic_server):
    observations = table_observations(); asset_header, header, _, _, row, region = observations
    foreign = plan(region, header, asset_header, row); foreign["plans"][0]["region_observation_id"] = "other:region"
    AnthropicStubHandler.response_content = foreign
    with pytest.raises(MapperValidationError):
        OpenAICompatibleMapper(anthropic_config(anthropic_server)).map("d", observations)


@pytest.mark.parametrize(("stop_reason", "message"), [("refusal", "mapper_refusal"), ("max_tokens", "mapper_incomplete:max_output_tokens"), ("pause_turn", "mapper_status:pause_turn")])
def test_anthropic_mapper_unfinished_stop_reasons_are_unavailable(anthropic_server, stop_reason, message):
    AnthropicStubHandler.response_content = {"plans": []}; AnthropicStubHandler.stop_reason = stop_reason
    with pytest.raises(MapperUnavailableError, match=message):
        OpenAICompatibleMapper(anthropic_config(anthropic_server)).map("d", table_observations())


def test_anthropic_overload_is_retried_but_billing_error_is_not(provider_error_server):
    completed = {"type": "message", "content": [{"type": "text", "text": json.dumps({"plans": []})}], "stop_reason": "end_turn"}
    ProviderErrorHandler.responses = [(529, {"type": "error", "error": {"type": "overloaded_error", "message": "x"}}, {"Retry-After": "0"}), (200, completed, {})]
    mapper = OpenAICompatibleMapper(anthropic_config(provider_error_server))
    assert mapper.map("d", table_observations(), transient_retries=2) == ()
    assert ProviderErrorHandler.calls == 2
    ProviderErrorHandler.calls = 0
    ProviderErrorHandler.responses = [(402, {"type": "error", "error": {"type": "billing_error", "message": "card details"}}, {"request-id": "req_9"})]
    with pytest.raises(MapperUnavailableError) as caught:
        mapper.map("d", table_observations(), transient_retries=2)
    assert ProviderErrorHandler.calls == 1 and "card details" not in str(caught.value)
    assert mapper.last_provider_diagnostic == {"status": 402, "message": "mapper provider returned HTTP 402", "type": "billing_error", "request_id": "req_9"}


def test_mapper_provider_defaults_fingerprints_and_credentials(monkeypatch):
    openai = OpenAICompatibleMapper(MapperConfig(True, model="m"))
    assert "provider" not in openai.fingerprint() and openai.endpoint == "https://api.openai.com/v1/responses"
    monkeypatch.delenv("ANTHROPIC_MODEL", raising=False); monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    claude = OpenAICompatibleMapper(MapperConfig(True, provider="anthropic"))
    assert claude.endpoint == "https://api.anthropic.com/v1/messages" and claude.model == "claude-opus-5-5"
    assert claude.fingerprint()["provider"] == "anthropic"
    with pytest.raises(MapperUnavailableError, match="ANTHROPIC_API_KEY is not configured"):
        claude.map("d", table_observations())
    with pytest.raises(MapperValidationError, match="unknown mapper provider"):
        OpenAICompatibleMapper(MapperConfig(True, provider="other")).map("d", table_observations())
