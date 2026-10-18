from __future__ import annotations

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from concurrent.futures import ThreadPoolExecutor
import json
from threading import Thread
import time

import pytest

from carthagent.ocr_azure import (
    AzureDocumentIntelligenceOCR,
    AzureOCRAnalysisError,
    AzureOCRConfig,
    AzureOCRConfigError,
    AzureOCRResourceLimitError,
    AzureOCRResponseError,
    AzureOCRUnavailableError,
)


API_VERSION = "test-api-version"
KEY_ENV = "CARTHAGENT_TEST_AZURE_KEY"


def succeeded_payload() -> dict[str, object]:
    return {
        "status": "succeeded",
        "analyzeResult": {
            "apiVersion": API_VERSION,
            "modelId": "prebuilt-layout",
            "stringIndexType": "unicodeCodePoint",
            "content": "Asset A 100",
            "pages": [
                {
                    "pageNumber": 1,
                    "width": 8.5,
                    "height": 11,
                    "unit": "inch",
                    "angle": 90,
                    "words": [
                        {
                            "content": "Asset",
                            "polygon": [0.1, 0.2, 1.1, 0.2, 1.1, 0.4, 0.1, 0.4],
                            "confidence": 0.991,
                            "span": {"offset": 0, "length": 5},
                        }
                    ],
                }
            ],
            "tables": [
                {
                    "rowCount": 2,
                    "columnCount": 2,
                    "boundingRegions": [
                        {"pageNumber": 1, "polygon": [0, 0, 2, 0, 2, 2, 0, 2]}
                    ],
                    "cells": [
                        {
                            "rowIndex": 0,
                            "columnIndex": 0,
                            "rowSpan": 1,
                            "columnSpan": 2,
                            "kind": "columnHeader",
                            "content": "Asset",
                            "spans": [{"offset": 0, "length": 5}],
                            "boundingRegions": [
                                {"pageNumber": 1, "polygon": [0, 0, 2, 0, 2, 1, 0, 1]}
                            ],
                        }
                    ],
                }
            ],
        },
    }


class AzureStub(BaseHTTPRequestHandler):
    posts = 0
    gets = 0
    received_body = b""
    received_key_headers: list[str | None] = []
    post_failures = 0
    poll_payloads: list[dict[str, object] | bytes] = []
    operation_override: str | None = None
    post_path = ""
    post_content_type = ""

    @classmethod
    def reset(cls) -> None:
        cls.posts = 0
        cls.gets = 0
        cls.received_body = b""
        cls.received_key_headers = []
        cls.post_failures = 0
        cls.poll_payloads = [succeeded_payload()]
        cls.operation_override = None
        cls.post_path = ""
        cls.post_content_type = ""

    def do_POST(self) -> None:
        type(self).posts += 1
        type(self).post_path = self.path
        type(self).post_content_type = self.headers.get("Content-Type", "")
        type(self).received_key_headers.append(self.headers.get("Ocp-Apim-Subscription-Key"))
        type(self).received_body = self.rfile.read(int(self.headers["Content-Length"]))
        if type(self).posts <= type(self).post_failures:
            self.send_error(500)
            return
        operation = type(self).operation_override or f"http://127.0.0.1:{self.server.server_port}/operations/abc"
        self.send_response(202)
        self.send_header("Operation-Location", operation)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:
        type(self).gets += 1
        type(self).received_key_headers.append(self.headers.get("Ocp-Apim-Subscription-Key"))
        index = min(type(self).gets - 1, len(type(self).poll_payloads) - 1)
        payload = type(self).poll_payloads[index]
        body = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args: object) -> None:
        return


@pytest.fixture
def azure_server():
    AzureStub.reset()
    server = ThreadingHTTPServer(("127.0.0.1", 0), AzureStub)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{server.server_port}"
    server.shutdown()
    thread.join()


def config(endpoint: str, **changes: object) -> AzureOCRConfig:
    values: dict[str, object] = {
        "endpoint": endpoint,
        "api_version": API_VERSION,
        "key_env": KEY_ENV,
        "timeout_seconds": 2,
        "transient_retries": 2,
        "retry_delay_seconds": 0,
        "max_input_bytes": 1024,
        "max_response_bytes": 20_000,
        "poll_interval_seconds": 0,
    }
    values.update(changes)
    return AzureOCRConfig(**values)  # type: ignore[arg-type]


def test_layout_request_and_provider_neutral_result(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "test-secret")
    document = b"%PDF-synthetic-private-bytes"
    result = AzureDocumentIntelligenceOCR(config(azure_server)).analyze(document)

    assert AzureStub.posts == 1 and AzureStub.gets == 1
    assert AzureStub.received_body == document
    assert AzureStub.post_path == "/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=test-api-version&stringIndexType=unicodeCodePoint"
    assert AzureStub.post_content_type == "application/octet-stream"
    assert AzureStub.received_key_headers == ["test-secret", "test-secret"]
    assert result.provider == "azure_document_intelligence"
    assert result.api_version == API_VERSION and result.model_id == "prebuilt-layout"
    assert result.string_index_type == "unicodeCodePoint"
    assert result.pages[0].unit == "inch" and result.pages[0].angle == 90
    assert result.pages[0].words[0].polygon == (0.1, 0.2, 1.1, 0.2, 1.1, 0.4, 0.1, 0.4)
    assert result.pages[0].words[0].confidence == 0.991
    assert result.pages[0].words[0].span.offset == 0
    cell = result.tables[0].cells[0]
    assert (cell.row_index, cell.column_index, cell.row_span, cell.column_span) == (0, 0, 1, 2)
    assert cell.spans[0].length == 5 and cell.bounding_regions[0].page_number == 1


def test_running_poll_and_transient_post_retry(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    AzureStub.post_failures = 1
    AzureStub.poll_payloads = [{"status": "running"}, succeeded_payload()]
    result = AzureDocumentIntelligenceOCR(config(azure_server)).analyze(b"pdf")
    assert result.content == "Asset A 100"
    assert AzureStub.posts == 2 and AzureStub.gets == 2


def test_missing_key_size_and_deadline_fail_before_upload(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.delenv(KEY_ENV, raising=False)
    adapter = AzureDocumentIntelligenceOCR(config(azure_server, max_input_bytes=3))
    with pytest.raises(AzureOCRResourceLimitError, match="empty_input"):
        adapter.analyze(b"")
    with pytest.raises(AzureOCRResourceLimitError, match="input_size_exceeded"):
        adapter.analyze(b"four")
    with pytest.raises(AzureOCRUnavailableError, match=KEY_ENV):
        adapter.analyze(b"one")
    monkeypatch.setenv(KEY_ENV, "secret")
    with pytest.raises(AzureOCRUnavailableError, match="deadline_exceeded"):
        adapter.analyze(b"one", deadline=time.monotonic() - 1)
    assert AzureStub.posts == 0


def test_same_adapter_supports_concurrent_documents_and_unicode(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = succeeded_payload()
    payload["analyzeResult"]["content"] = "Anlage Ä 😀"
    payload["analyzeResult"]["pages"][0]["words"][0]["span"] = {"offset": 0, "length": 6}
    AzureStub.poll_payloads = [payload]
    adapter = AzureDocumentIntelligenceOCR(config(azure_server))
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = tuple(executor.map(adapter.analyze, (b"one", b"two", b"three", b"four")))
    assert [result.content for result in results] == ["Anlage Ä 😀"] * 4
    assert AzureStub.posts == 4 and AzureStub.gets == 4


def test_operation_location_cannot_exfiltrate_key(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "do-not-leak")
    AzureStub.operation_override = "http://example.invalid/steal"
    with pytest.raises(AzureOCRResponseError, match="origin_mismatch"):
        AzureDocumentIntelligenceOCR(config(azure_server)).analyze(b"private")
    assert AzureStub.posts == 1 and AzureStub.gets == 0
    assert "do-not-leak" not in repr(config(azure_server))


@pytest.mark.parametrize(
    "mutation,error",
    [
        (lambda p: p["analyzeResult"].__setitem__("modelId", "other"), "model_id_mismatch"),
        (lambda p: p["analyzeResult"].__setitem__("stringIndexType", "utf16CodeUnit"), "string_index_type_mismatch"),
        (lambda p: p["analyzeResult"]["pages"][0].__setitem__("unit", "points"), "invalid_page_unit"),
        (lambda p: p["analyzeResult"]["pages"][0]["words"][0].__setitem__("confidence", 2), "confidence_out_of_range"),
        (lambda p: p["analyzeResult"]["tables"][0]["cells"][0]["boundingRegions"][0].__setitem__("pageNumber", 99), "unknown_page"),
        (lambda p: p["analyzeResult"]["tables"][0]["cells"][0].__setitem__("rowIndex", 2), "outside_table_dimensions"),
        (lambda p: p["analyzeResult"]["pages"][0]["words"][0]["span"].__setitem__("length", 999), "text_span_out_of_bounds"),
    ],
)
def test_strict_response_validation(azure_server: str, monkeypatch: pytest.MonkeyPatch, mutation, error: str):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = succeeded_payload()
    mutation(payload)
    AzureStub.poll_payloads = [payload]
    with pytest.raises(AzureOCRResponseError, match=error):
        AzureDocumentIntelligenceOCR(config(azure_server)).analyze(b"pdf")


def test_invalid_json_duplicate_keys_response_bound_and_failed_job_are_explicit(azure_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret-value")
    adapter = AzureDocumentIntelligenceOCR(config(azure_server))
    AzureStub.poll_payloads = [b'{"status":"running","status":"succeeded"}']
    with pytest.raises(AzureOCRResponseError, match="duplicate_json_key"):
        adapter.analyze(b"document-contents")

    AzureStub.reset()
    AzureStub.poll_payloads = [{"status": "failed", "error": {"code": "InvalidContent", "message": "document-contents"}}]
    with pytest.raises(AzureOCRAnalysisError, match="InvalidContent") as failure:
        adapter.analyze(b"document-contents")
    assert "document-contents" not in str(failure.value) and "secret-value" not in str(failure.value)

    AzureStub.reset()
    AzureStub.poll_payloads = [succeeded_payload()]
    with pytest.raises(AzureOCRResourceLimitError, match="response_size_exceeded"):
        AzureDocumentIntelligenceOCR(config(azure_server, max_response_bytes=10)).analyze(b"pdf")


def test_config_rejects_secret_bearing_or_unbounded_endpoints():
    with pytest.raises(AzureOCRConfigError, match="endpoint_invalid"):
        config("https://user:secret@example.com")
    with pytest.raises(AzureOCRConfigError, match="endpoint_invalid"):
        config("https://example.com?api-key=secret")
    with pytest.raises(AzureOCRConfigError, match="requires_https"):
        config("http://example.com")
    with pytest.raises(AzureOCRConfigError, match="size_bounds"):
        config("https://example.com", max_response_bytes=0)
    with pytest.raises(AzureOCRConfigError, match="retries"):
        config("https://example.com", transient_retries=1.5)
    with pytest.raises(AzureOCRConfigError, match="timeout"):
        config("https://example.com", timeout_seconds=True)
    with pytest.raises(AzureOCRConfigError, match="retry_delay"):
        config("https://example.com", retry_delay_seconds=-1)
