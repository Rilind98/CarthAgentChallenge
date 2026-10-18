from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from threading import Lock, Thread
import time

import pytest

from carthagent.ocr_mistral import (
    MISTRAL_OCR_CONFIDENCE_GRANULARITY,
    MISTRAL_OCR_QUALITY_MODEL,
    MISTRAL_OCR_TABLE_FORMAT,
    MistralDocumentOCR,
    MistralOCRAuthenticationError,
    MistralOCRConfig,
    MistralOCRConfigError,
    MistralOCRDeadlineError,
    MistralOCRProviderError,
    MistralOCRResourceLimitError,
    MistralOCRResponseError,
    MistralOCRUnavailableError,
    _decode_json,
    _parse_result,
)


KEY_ENV = "CARTHAGENT_TEST_MISTRAL_KEY"


def successful_payload() -> dict[str, object]:
    markdown = "Anlage Ä 210003"
    table_html = "<table><tr><th>ID</th></tr><tr><td>210003</table>"
    return {
        "model": "mistral-ocr-4.1",
        "pages": [
            {
                "index": 0,
                "markdown": markdown,
                "dimensions": {"dpi": 200, "width": 1700, "height": 2200},
                "images": [
                    {
                        "id": "img-0",
                        "top_left_x": None,
                        "top_left_y": None,
                        "bottom_right_x": None,
                        "bottom_right_y": None,
                        "image_base64": None,
                        "image_annotation": "Diagramm Ä",
                    }
                ],
                "tables": [
                    {
                        "id": "tbl-0",
                        "content": table_html,
                        "format": "html",
                        "word_confidence_scores": [
                            {"text": "210003", "confidence": 0.96, "start_index": 35}
                        ],
                    }
                ],
                "blocks": [
                    {
                        "type": "text",
                        "content": "Anlage Ä",
                        "top_left_x": 100,
                        "top_left_y": 200,
                        "bottom_right_x": 600,
                        "bottom_right_y": 300,
                    },
                    {
                        "type": "table",
                        "content": table_html,
                        "top_left_x": 100,
                        "top_left_y": 400,
                        "bottom_right_x": 1600,
                        "bottom_right_y": 1000,
                        "table_id": "tbl-0",
                    },
                    {"type": "image", "content": "", "image_id": "img-0"},
                ],
                "confidence_scores": {
                    "average_page_confidence_score": 0.97,
                    "minimum_page_confidence_score": 0.91,
                    "word_confidence_scores": [
                        {"text": "Anlage", "confidence": 0.99, "start_index": 0},
                        {"text": "Ä", "confidence": 0.95, "start_index": 7},
                        {"text": "210003", "confidence": 0.96, "start_index": 9},
                    ],
                },
            }
        ],
        "usage_info": {"pages_processed": 1, "doc_size_bytes": 20},
    }


class MistralStub(BaseHTTPRequestHandler):
    lock = Lock()
    requests = 0
    received: list[tuple[str, str | None, bytes]] = []
    responses: list[tuple[int, dict[str, str], bytes, float]] = []

    @classmethod
    def reset(cls) -> None:
        with cls.lock:
            cls.requests = 0
            cls.received = []
            cls.responses = [(200, {"Content-Type": "application/json"}, json.dumps(successful_payload()).encode(), 0)]

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        with type(self).lock:
            index = type(self).requests
            type(self).requests += 1
            type(self).received.append((self.path, self.headers.get("Authorization"), body))
            response = type(self).responses[min(index, len(type(self).responses) - 1)]
        status, headers, response_body, delay = response
        if delay:
            time.sleep(delay)
        self.send_response(status)
        for key, value in headers.items():
            self.send_header(key, value)
        self.send_header("Content-Length", str(len(response_body)))
        self.end_headers()
        try:
            self.wfile.write(response_body)
        except BrokenPipeError:
            pass

    def log_message(self, format: str, *args: object) -> None:
        return


@pytest.fixture
def mistral_server():
    MistralStub.reset()
    server = ThreadingHTTPServer(("127.0.0.1", 0), MistralStub)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{server.server_port}/v1/ocr"
    server.shutdown()
    thread.join()


def config(endpoint: str, **changes: object) -> MistralOCRConfig:
    values: dict[str, object] = {
        "endpoint": endpoint,
        "model": MISTRAL_OCR_QUALITY_MODEL,
        "key_env": KEY_ENV,
        "timeout_seconds": 2,
        "transient_retries": 2,
        "retry_delay_seconds": 0,
        "max_retry_after_seconds": 1,
        "max_input_bytes": 1024,
        "max_response_bytes": 100_000,
    }
    values.update(changes)
    return MistralOCRConfig(**values)  # type: ignore[arg-type]


def set_response(payload: dict[str, object] | bytes, *, status: int = 200, headers: dict[str, str] | None = None, delay: float = 0) -> None:
    body = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
    MistralStub.responses = [(status, headers or {"Content-Type": "application/json"}, body, delay)]


def test_request_contract_and_provider_neutral_evidence(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "test-secret")
    result = MistralDocumentOCR(config(mistral_server)).analyze(b"%PDF-synthetic")

    assert MistralStub.requests == 1
    path, authorization, raw_request = MistralStub.received[0]
    request = json.loads(raw_request)
    assert path == "/v1/ocr" and authorization == "Bearer test-secret"
    assert request["model"] == MISTRAL_OCR_QUALITY_MODEL
    assert request["table_format"] == MISTRAL_OCR_TABLE_FORMAT
    assert request["include_image_base64"] is False
    assert request["include_blocks"] is True
    assert request["confidence_scores_granularity"] == MISTRAL_OCR_CONFIDENCE_GRANULARITY
    assert request["document"]["type"] == "document_url"
    assert request["document"]["document_url"].startswith("data:application/pdf;base64,")

    assert result.provider == "mistral_document_ai"
    assert result.requested_model == "mistral-ocr-latest"
    assert result.provider_model == "mistral-ocr-4.1"
    assert result.usage.pages_processed == 1
    page = result.pages[0]
    assert page.index == 0 and page.markdown == "Anlage Ä 210003"
    assert page.dimensions is not None
    assert page.dimensions.pixel_width == 1700
    assert page.dimensions.pdf_width_points == pytest.approx(612)
    assert page.confidence is not None and page.confidence.minimum == 0.91
    assert [item.text for item in page.recognition_evidence] == ["Anlage", "Ä", "210003"]
    assert all(item.geometry is None and item.evidence_kind == "recognition_only" for item in page.recognition_evidence)
    assert all(item.content_match_kind == "exact" and item.content_span_text == item.text for item in page.recognition_evidence)
    assert page.blocks[0].geometry is not None
    assert page.blocks[0].geometry.provider_bbox == (100.0, 200.0, 600.0, 300.0)
    assert page.blocks[0].geometry.normalized_bbox == pytest.approx((100 / 1700, 200 / 2200, 600 / 1700, 300 / 2200))
    assert page.blocks[0].geometry.pdf_bbox == pytest.approx((36, 72, 216, 108))
    assert page.blocks[0].geometry.coordinate_authority == "candidate_only"
    assert page.blocks[0].geometry.requires_independent_anchor is True
    assert page.blocks[2].geometry is None
    assert page.images[0].geometry is None


def test_malformed_table_html_is_preserved_and_never_promoted_to_cells(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    result = MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")
    table = result.pages[0].tables[0]
    assert table.content == "<table><tr><th>ID</th></tr><tr><td>210003</table>"
    assert table.format == "html"
    assert table.cell_structure_authoritative is False
    assert table.content_trust == "untrusted_provider_output"
    assert table.recognition_evidence[0].geometry is None


def test_page_confidence_accepts_only_provider_table_placeholder_rebasing(
    mistral_server: str, monkeypatch: pytest.MonkeyPatch
):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = successful_payload()
    page = payload["pages"][0]
    page["markdown"] = "Intro\n\n[tbl-10.html](tbl-10.html)\n\nEnd"
    page["tables"][0]["id"] = "tbl-10.html"
    page["blocks"][1]["table_id"] = "tbl-10.html"
    page["confidence_scores"]["word_confidence_scores"] = [
        {"text": "Intro", "confidence": 0.99, "start_index": 0},
        {"text": "[tbl-0.html](tbl-0.html)", "confidence": 0.95, "start_index": 7},
        {"text": "End", "confidence": 0.96, "start_index": 35},
    ]
    set_response(payload)

    result = MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")
    link = result.pages[0].recognition_evidence[1]
    assert link.text == "[tbl-0.html](tbl-0.html)"
    assert link.content_span_text == "[tbl-10.html](tbl-10.html)"
    assert link.content_match_kind == "rebased_table_placeholder"
    assert link.geometry is None


def test_non_placeholder_confidence_mismatch_remains_invalid(
    mistral_server: str, monkeypatch: pytest.MonkeyPatch
):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = successful_payload()
    payload["pages"][0]["confidence_scores"]["word_confidence_scores"][0]["text"] = "Wrong"
    set_response(payload)
    with pytest.raises(MistralOCRResponseError, match="word_confidence_text_mismatch"):
        MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")

    payload = successful_payload()
    page = payload["pages"][0]
    page["markdown"] = "[tbl-9.html](tbl-9.html)"
    page["confidence_scores"]["word_confidence_scores"] = [
        {"text": "[tbl-0.html](tbl-0.html)", "confidence": 0.95, "start_index": 0}
    ]
    set_response(payload)
    with pytest.raises(MistralOCRResponseError, match="word_confidence_text_mismatch"):
        MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")


def test_saved_real_mistral_responses_replay_read_only():
    response_dir = Path("hackathon-local/mistral-benchmark-v1")
    paths = tuple(
        response_dir / name
        for name in (
            "f00_demo_03.response.json",
            "f01_net_spiegel_05.response.json",
            "f03_block_spiegel_03.response.json",
            "f07_sap_class_summary_03.response.json",
        )
    )
    if not all(path.is_file() for path in paths):
        pytest.skip("local Mistral benchmark responses are unavailable")
    for path in paths:
        result = _parse_result(
            _decode_json(path.read_bytes()),
            requested_model=MISTRAL_OCR_QUALITY_MODEL,
        )
        assert result.pages


def test_saved_f04_response_retains_strict_duplicate_table_reference_rejection():
    path = Path("hackathon-local/mistral-benchmark-v1/f04_verzeichnis_01.response.json")
    if not path.is_file():
        pytest.skip("local f04 Mistral response is unavailable")
    with pytest.raises(MistralOCRResponseError, match="invalid_or_duplicate_table_reference"):
        _parse_result(
            _decode_json(path.read_bytes()),
            requested_model=MISTRAL_OCR_QUALITY_MODEL,
        )


def test_optional_polygon_and_missing_geometry_are_preserved_without_fabrication(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = successful_payload()
    blocks = payload["pages"][0]["blocks"]
    blocks[0].pop("top_left_x")
    blocks[0].pop("top_left_y")
    blocks[0].pop("bottom_right_x")
    blocks[0].pop("bottom_right_y")
    blocks[0]["polygon"] = [100, 200, 600, 200, 600, 300, 100, 300]
    set_response(payload)

    result = MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")
    geometry = result.pages[0].blocks[0].geometry
    assert geometry is not None
    assert geometry.provider_polygon == (100.0, 200.0, 600.0, 200.0, 600.0, 300.0, 100.0, 300.0)
    assert geometry.pdf_polygon == pytest.approx((36, 72, 216, 72, 216, 108, 36, 108))
    assert result.pages[0].blocks[2].geometry is None


def test_missing_auth_and_input_limits_fail_before_network(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.delenv(KEY_ENV, raising=False)
    adapter = MistralDocumentOCR(config(mistral_server, max_input_bytes=3))
    with pytest.raises(MistralOCRResourceLimitError, match="empty_input"):
        adapter.analyze(b"")
    with pytest.raises(MistralOCRResourceLimitError, match="input_size_exceeded"):
        adapter.analyze(b"four")
    with pytest.raises(MistralOCRAuthenticationError, match="key_env_unavailable"):
        adapter.analyze(b"pdf")
    assert MistralStub.requests == 0


def test_endpoint_origin_policy_and_redirect_rejection(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    with pytest.raises(MistralOCRConfigError, match="requires_https_or_loopback"):
        config("http://example.com/v1/ocr")
    with pytest.raises(MistralOCRConfigError, match="endpoint_invalid"):
        config("https://user:secret@example.com/v1/ocr")
    with pytest.raises(MistralOCRConfigError, match="endpoint_invalid"):
        config("https://example.com/v1/ocr?token=secret")
    with pytest.raises(MistralOCRConfigError, match="origin_not_allowed"):
        config("https://example.com/v1/ocr")
    with pytest.raises(MistralOCRConfigError, match="origin_not_allowed"):
        config("https://api.mistral.ai:444/v1/ocr")

    monkeypatch.setenv(KEY_ENV, "secret")
    set_response(b"", status=302, headers={"Location": "https://example.invalid/steal"})
    with pytest.raises(MistralOCRProviderError) as failure:
        MistralDocumentOCR(config(mistral_server)).analyze(b"private-document")
    assert failure.value.status_code == 302
    assert MistralStub.requests == 1


def test_terminal_400_has_safe_code_without_provider_message_or_document(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret-value")
    set_response(
        {"error": {"code": "invalid_document", "message": "private-document secret-value"}},
        status=400,
    )
    with pytest.raises(MistralOCRProviderError) as failure:
        MistralDocumentOCR(config(mistral_server)).analyze(b"private-document")
    assert failure.value.status_code == 400
    assert failure.value.provider_code == "invalid_document"
    assert failure.value.retryable is False
    assert "private-document" not in str(failure.value)
    assert "secret-value" not in str(failure.value)
    assert MistralStub.requests == 1


def test_transient_retry_obeys_retry_after_and_is_bounded(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    success = json.dumps(successful_payload()).encode()
    MistralStub.responses = [
        (429, {"Content-Type": "application/json", "Retry-After": "0"}, b'{"error":{"code":"rate_limit"}}', 0),
        (200, {"Content-Type": "application/json"}, success, 0),
    ]
    result = MistralDocumentOCR(config(mistral_server, transient_retries=1)).analyze(b"pdf")
    assert result.pages[0].index == 0 and MistralStub.requests == 2

    MistralStub.reset()
    MistralStub.responses = [
        (503, {"Content-Type": "application/json", "Retry-After": "30"}, b'{"error":{"code":"overloaded"}}', 0)
    ]
    with pytest.raises(MistralOCRProviderError) as failure:
        MistralDocumentOCR(config(mistral_server, transient_retries=2, max_retry_after_seconds=1)).analyze(b"pdf")
    assert failure.value.provider_code == "overloaded"
    assert MistralStub.requests == 1


def test_deadline_and_transport_timeout_are_typed(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    adapter = MistralDocumentOCR(config(mistral_server))
    with pytest.raises(MistralOCRDeadlineError, match="deadline_exceeded"):
        adapter.analyze(b"pdf", deadline=time.monotonic() - 1)
    assert MistralStub.requests == 0

    set_response(successful_payload(), delay=0.1)
    with pytest.raises(MistralOCRUnavailableError, match="transport_unavailable"):
        MistralDocumentOCR(config(mistral_server, timeout_seconds=0.02, transient_retries=0)).analyze(b"pdf")


def test_response_size_is_bounded(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    with pytest.raises(MistralOCRResourceLimitError, match="response_size_exceeded"):
        MistralDocumentOCR(config(mistral_server, max_response_bytes=10)).analyze(b"pdf")


@pytest.mark.parametrize(
    "body,error",
    [
        (b"{", "invalid_json"),
        (b'{"model":"x","pages":[],"pages":[],"usage_info":{"pages_processed":0}}', "duplicate_json_key"),
        (b'{"model":"x","pages":[],"usage_info":{"pages_processed":NaN}}', "invalid_json"),
    ],
)
def test_malformed_duplicate_and_nonfinite_json_are_rejected(
    mistral_server: str, monkeypatch: pytest.MonkeyPatch, body: bytes, error: str
):
    monkeypatch.setenv(KEY_ENV, "secret")
    set_response(body)
    with pytest.raises(MistralOCRResponseError, match=error):
        MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")


@pytest.mark.parametrize(
    "mutation,error",
    [
        (lambda p: p["pages"].append(deepcopy(p["pages"][0])), "duplicate_page_index"),
        (lambda p: p["pages"][0]["tables"].append(deepcopy(p["pages"][0]["tables"][0])), "duplicate_or_empty_table_id"),
        (lambda p: p["pages"][0]["tables"][0].__setitem__("format", "markdown"), "unexpected_table_format"),
        (lambda p: p["pages"][0]["blocks"][0].__setitem__("bottom_right_x", 9999), "geometry_out_of_bounds"),
        (lambda p: p["pages"][0]["confidence_scores"].__setitem__("average_page_confidence_score", 2), "confidence_out_of_range"),
        (lambda p: p["pages"][0]["confidence_scores"]["word_confidence_scores"].append(deepcopy(p["pages"][0]["confidence_scores"]["word_confidence_scores"][0])), "duplicate_word_confidence"),
        (lambda p: p["pages"][0].__setitem__("dimensions", None), "usage_page_count_mismatch"),
    ],
)
def test_strict_response_validation(
    mistral_server: str, monkeypatch: pytest.MonkeyPatch, mutation, error: str
):
    monkeypatch.setenv(KEY_ENV, "secret")
    payload = successful_payload()
    if error == "usage_page_count_mismatch":
        payload["usage_info"]["pages_processed"] = 2
    else:
        mutation(payload)
    set_response(payload)
    with pytest.raises(MistralOCRResponseError, match=error):
        MistralDocumentOCR(config(mistral_server)).analyze(b"pdf")


def test_same_adapter_is_safe_for_concurrent_documents_and_unicode(mistral_server: str, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv(KEY_ENV, "secret")
    adapter = MistralDocumentOCR(config(mistral_server))
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = tuple(executor.map(adapter.analyze, (b"one", b"two", b"three", b"four")))
    assert [result.pages[0].markdown for result in results] == ["Anlage Ä 210003"] * 4
    assert MistralStub.requests == 4


def test_config_numeric_bounds_are_strict():
    with pytest.raises(MistralOCRConfigError, match="timeout"):
        config("https://api.mistral.ai/v1/ocr", timeout_seconds=True)
    with pytest.raises(MistralOCRConfigError, match="retries"):
        config("https://api.mistral.ai/v1/ocr", transient_retries=1.5)
    with pytest.raises(MistralOCRConfigError, match="retry_delay"):
        config("https://api.mistral.ai/v1/ocr", retry_delay_seconds=-1)
    with pytest.raises(MistralOCRConfigError, match="size_bounds"):
        config("https://api.mistral.ai/v1/ocr", max_response_bytes=0)
