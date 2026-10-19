from __future__ import annotations

from dataclasses import dataclass
import json
import math
import os
import socket
import time
from typing import Any, Mapping, Sequence
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlparse
from urllib.request import HTTPRedirectHandler, Request, build_opener


AZURE_LAYOUT_MODEL = "prebuilt-layout"
AZURE_STRING_INDEX_TYPE = "unicodeCodePoint"


class AzureOCRError(RuntimeError):
    """Base class for bounded, non-destructive Azure OCR failures."""


class AzureOCRConfigError(AzureOCRError):
    pass


class AzureOCRUnavailableError(AzureOCRError):
    pass


class AzureOCRResourceLimitError(AzureOCRError):
    pass


class AzureOCRResponseError(AzureOCRError):
    pass


class AzureOCRAnalysisError(AzureOCRError):
    pass


@dataclass(frozen=True, slots=True)
class AzureOCRConfig:
    endpoint: str
    api_version: str
    key_env: str
    timeout_seconds: float
    transient_retries: int
    retry_delay_seconds: float
    max_input_bytes: int
    max_response_bytes: int
    poll_interval_seconds: float

    def __post_init__(self) -> None:
        parsed = urlparse(self.endpoint)
        loopback = parsed.hostname in {"127.0.0.1", "::1", "localhost"}
        if parsed.scheme not in ({"https"} if not loopback else {"http", "https"}):
            raise AzureOCRConfigError("azure_ocr_endpoint_requires_https")
        if not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
            raise AzureOCRConfigError("azure_ocr_endpoint_invalid")
        if not isinstance(self.api_version, str) or not self.api_version.strip() or not isinstance(self.key_env, str) or not self.key_env.strip():
            raise AzureOCRConfigError("azure_ocr_api_version_and_key_env_required")
        if isinstance(self.timeout_seconds, bool) or not isinstance(self.timeout_seconds, (int, float)) or not math.isfinite(self.timeout_seconds) or self.timeout_seconds <= 0:
            raise AzureOCRConfigError("azure_ocr_timeout_must_be_positive")
        if isinstance(self.transient_retries, bool) or not isinstance(self.transient_retries, int) or self.transient_retries < 0:
            raise AzureOCRConfigError("azure_ocr_retries_must_be_nonnegative")
        if isinstance(self.retry_delay_seconds, bool) or not isinstance(self.retry_delay_seconds, (int, float)) or not math.isfinite(self.retry_delay_seconds) or self.retry_delay_seconds < 0:
            raise AzureOCRConfigError("azure_ocr_retry_delay_must_be_nonnegative")
        if any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in (self.max_input_bytes, self.max_response_bytes)):
            raise AzureOCRConfigError("azure_ocr_size_bounds_must_be_positive")
        if isinstance(self.poll_interval_seconds, bool) or not isinstance(self.poll_interval_seconds, (int, float)) or not math.isfinite(self.poll_interval_seconds) or self.poll_interval_seconds < 0:
            raise AzureOCRConfigError("azure_ocr_poll_interval_must_be_nonnegative")


@dataclass(frozen=True, slots=True)
class OCRTextSpan:
    offset: int
    length: int


@dataclass(frozen=True, slots=True)
class OCRBoundingRegion:
    page_number: int
    polygon: tuple[float, ...]


@dataclass(frozen=True, slots=True)
class OCRWord:
    text: str
    polygon: tuple[float, ...]
    confidence: float
    span: OCRTextSpan


@dataclass(frozen=True, slots=True)
class OCRPage:
    page_number: int
    width: float
    height: float
    unit: str
    angle: float | None
    words: tuple[OCRWord, ...]


@dataclass(frozen=True, slots=True)
class OCRTableCell:
    row_index: int
    column_index: int
    row_span: int | None
    column_span: int | None
    kind: str | None
    content: str
    spans: tuple[OCRTextSpan, ...]
    bounding_regions: tuple[OCRBoundingRegion, ...]


@dataclass(frozen=True, slots=True)
class OCRTable:
    row_count: int
    column_count: int
    cells: tuple[OCRTableCell, ...]
    bounding_regions: tuple[OCRBoundingRegion, ...]


@dataclass(frozen=True, slots=True)
class OCRResult:
    provider: str
    model_id: str
    api_version: str
    string_index_type: str
    content: str
    pages: tuple[OCRPage, ...]
    tables: tuple[OCRTable, ...]


class _RejectRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):  # type: ignore[no-untyped-def]
        return None


class AzureDocumentIntelligenceOCR:
    """Strict Azure Document Intelligence adapter with no implicit credentials."""

    def __init__(self, config: AzureOCRConfig):
        self.config = config
        self._opener = build_opener(_RejectRedirects())

    def analyze(self, document_bytes: bytes, *, deadline: float | None = None, pages: Sequence[int] | None = None) -> OCRResult:
        """Analyse the document; ``pages`` (1-based) limits analysis and billing to those pages."""
        if not isinstance(document_bytes, bytes):
            raise TypeError("document_bytes must be bytes")
        page_query = _page_ranges(pages) if pages is not None else None
        if not document_bytes:
            raise AzureOCRResourceLimitError("azure_ocr_empty_input")
        if len(document_bytes) > self.config.max_input_bytes:
            raise AzureOCRResourceLimitError("azure_ocr_input_size_exceeded")
        api_key = os.environ.get(self.config.key_env)
        if not api_key:
            raise AzureOCRUnavailableError(f"azure_ocr_key_env_unavailable:{self.config.key_env}")
        local_deadline = time.monotonic() + self.config.timeout_seconds
        if deadline is not None:
            local_deadline = min(local_deadline, deadline)

        endpoint = self.config.endpoint.rstrip("/")
        parameters = {"api-version": self.config.api_version, "stringIndexType": AZURE_STRING_INDEX_TYPE}
        if page_query is not None:
            parameters["pages"] = page_query
        query = urlencode(parameters, safe=",")
        analyze_url = f"{endpoint}/documentintelligence/documentModels/{AZURE_LAYOUT_MODEL}:analyze?{query}"
        headers = {
            "Content-Type": "application/octet-stream",
            "Ocp-Apim-Subscription-Key": api_key,
        }
        response = self._request(
            Request(analyze_url, data=document_bytes, headers=headers, method="POST"),
            deadline=local_deadline,
            expected_status=202,
        )
        operation_url = response.headers.get("Operation-Location")
        if not operation_url:
            raise AzureOCRResponseError("azure_ocr_missing_operation_location")
        self._validate_operation_url(operation_url)

        while True:
            poll = self._request(
                Request(operation_url, headers={"Ocp-Apim-Subscription-Key": api_key}, method="GET"),
                deadline=local_deadline,
                expected_status=200,
            )
            content_type = poll.headers.get("Content-Type", "").partition(";")[0].strip().lower()
            if content_type != "application/json":
                raise AzureOCRResponseError("azure_ocr_poll_response_not_json")
            payload = _decode_json(poll.body)
            status = payload.get("status") if isinstance(payload, dict) else None
            if status == "succeeded":
                return _parse_result(payload, expected_api_version=self.config.api_version)
            if status == "failed":
                error = payload.get("error")
                code = error.get("code") if isinstance(error, dict) else None
                safe_code = code if isinstance(code, str) and code else "unknown"
                raise AzureOCRAnalysisError(f"azure_ocr_analysis_failed:{safe_code}")
            if status not in {"notStarted", "running"}:
                raise AzureOCRResponseError("azure_ocr_invalid_operation_status")
            remaining = local_deadline - time.monotonic()
            if remaining <= 0:
                raise AzureOCRUnavailableError("azure_ocr_deadline_exceeded")
            time.sleep(min(self.config.poll_interval_seconds, remaining))

    def _request(self, request: Request, *, deadline: float, expected_status: int) -> _HTTPResponse:
        for attempt in range(self.config.transient_retries + 1):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise AzureOCRUnavailableError("azure_ocr_deadline_exceeded")
            try:
                with self._opener.open(request, timeout=remaining) as response:
                    status = response.getcode()
                    if status != expected_status:
                        raise AzureOCRUnavailableError(f"azure_ocr_unexpected_http_status:{status}")
                    body = _read_bounded(response, self.config.max_response_bytes)
                    return _HTTPResponse(response.headers, body)
            except HTTPError as exc:
                transient = exc.code in {408, 429} or 500 <= exc.code <= 599
                if not transient or attempt == self.config.transient_retries:
                    raise AzureOCRUnavailableError(f"azure_ocr_http_error:{exc.code}") from exc
                self._wait_before_retry(deadline, exc.headers.get("Retry-After"))
            except (URLError, TimeoutError, socket.timeout) as exc:
                if attempt == self.config.transient_retries:
                    raise AzureOCRUnavailableError(f"azure_ocr_request_failed:{type(exc).__name__}") from exc
                self._wait_before_retry(deadline, None)
        raise AzureOCRUnavailableError("azure_ocr_request_failed:no_response")

    def _wait_before_retry(self, deadline: float, retry_after: str | None) -> None:
        delay = self.config.retry_delay_seconds
        if retry_after is not None:
            try:
                parsed = float(retry_after)
                if math.isfinite(parsed) and parsed >= 0:
                    delay = parsed
            except ValueError:
                pass
        remaining = deadline - time.monotonic()
        if remaining <= 0 or delay >= remaining:
            raise AzureOCRUnavailableError("azure_ocr_deadline_exceeded")
        time.sleep(delay)

    def _validate_operation_url(self, operation_url: str) -> None:
        configured = urlparse(self.config.endpoint)
        operation = urlparse(operation_url)
        if (
            operation.scheme != configured.scheme
            or operation.netloc != configured.netloc
            or operation.username
            or operation.password
            or operation.fragment
        ):
            raise AzureOCRResponseError("azure_ocr_operation_location_origin_mismatch")


def _page_ranges(pages: Sequence[int]) -> str:
    """Format 1-based page numbers as Azure's ``pages`` query value, e.g. ``1-3,5``."""
    numbers = sorted(set(pages))
    if not numbers or any(isinstance(page, bool) or not isinstance(page, int) or page < 1 for page in numbers):
        raise ValueError("azure_ocr_pages_must_be_positive_integers")
    ranges: list[str] = []
    start = previous = numbers[0]
    for page in numbers[1:] + [None]:
        if page is not None and page == previous + 1:
            previous = page
            continue
        ranges.append(str(start) if start == previous else f"{start}-{previous}")
        if page is not None:
            start = previous = page
    return ",".join(ranges)


@dataclass(frozen=True, slots=True)
class _HTTPResponse:
    headers: Mapping[str, str]
    body: bytes


def _read_bounded(response: Any, maximum: int) -> bytes:
    body = response.read(maximum + 1)
    if len(body) > maximum:
        raise AzureOCRResourceLimitError("azure_ocr_response_size_exceeded")
    return body


def _decode_json(body: bytes) -> dict[str, Any]:
    def object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise AzureOCRResponseError("azure_ocr_duplicate_json_key")
            result[key] = value
        return result

    try:
        value = json.loads(
            body.decode("utf-8"),
            object_pairs_hook=object_pairs,
            parse_constant=lambda _: (_ for _ in ()).throw(ValueError("non-finite JSON number")),
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
        if isinstance(exc, AzureOCRResponseError):
            raise
        raise AzureOCRResponseError("azure_ocr_invalid_json") from exc
    if not isinstance(value, dict):
        raise AzureOCRResponseError("azure_ocr_response_must_be_object")
    return value


def _parse_result(payload: dict[str, Any], *, expected_api_version: str) -> OCRResult:
    result = _mapping(payload, "analyzeResult")
    api_version = _text(result, "apiVersion")
    model_id = _text(result, "modelId")
    string_index_type = _text(result, "stringIndexType")
    if api_version != expected_api_version:
        raise AzureOCRResponseError("azure_ocr_api_version_mismatch")
    if model_id != AZURE_LAYOUT_MODEL:
        raise AzureOCRResponseError("azure_ocr_model_id_mismatch")
    if string_index_type != AZURE_STRING_INDEX_TYPE:
        raise AzureOCRResponseError("azure_ocr_string_index_type_mismatch")
    content = _text(result, "content")
    pages_raw = _list(result, "pages")
    tables_raw = result.get("tables", [])
    if not isinstance(tables_raw, list):
        raise AzureOCRResponseError("azure_ocr_tables_must_be_array")

    pages: list[OCRPage] = []
    page_numbers: set[int] = set()
    for raw in pages_raw:
        page = _as_mapping(raw, "page")
        page_number = _positive_int(page, "pageNumber")
        if page_number in page_numbers:
            raise AzureOCRResponseError("azure_ocr_duplicate_page_number")
        page_numbers.add(page_number)
        unit = _text(page, "unit")
        if unit not in {"inch", "pixel"}:
            raise AzureOCRResponseError("azure_ocr_invalid_page_unit")
        angle_raw = page.get("angle")
        angle = None if angle_raw is None else _finite_number(angle_raw, "angle")
        words_raw = _list(page, "words")
        words = tuple(_parse_word(_as_mapping(word, "word"), len(content)) for word in words_raw)
        pages.append(
            OCRPage(
                page_number,
                _positive_number(page, "width"),
                _positive_number(page, "height"),
                unit,
                angle,
                words,
            )
        )

    tables = tuple(_parse_table(_as_mapping(table, "table"), page_numbers, len(content)) for table in tables_raw)
    return OCRResult("azure_document_intelligence", model_id, api_version, string_index_type, content, tuple(pages), tables)


def _parse_word(word: Mapping[str, Any], content_length: int) -> OCRWord:
    confidence = _finite_number(word.get("confidence"), "word.confidence")
    if not 0 <= confidence <= 1:
        raise AzureOCRResponseError("azure_ocr_word_confidence_out_of_range")
    span = _parse_span(_mapping(word, "span"))
    _validate_span_bounds(span, content_length)
    return OCRWord(
        _text(word, "content"),
        _polygon(word.get("polygon"), "word.polygon"),
        confidence,
        span,
    )


def _parse_table(table: Mapping[str, Any], pages: set[int], content_length: int) -> OCRTable:
    row_count = _positive_int(table, "rowCount")
    column_count = _positive_int(table, "columnCount")
    regions = _parse_regions(table.get("boundingRegions", []), pages)
    cells_raw = _list(table, "cells")
    cells: list[OCRTableCell] = []
    for raw in cells_raw:
        cell = _as_mapping(raw, "table.cell")
        row_span = _optional_positive_int(cell, "rowSpan")
        column_span = _optional_positive_int(cell, "columnSpan")
        kind = cell.get("kind")
        if kind is not None and (not isinstance(kind, str) or not kind):
            raise AzureOCRResponseError("azure_ocr_invalid_cell_kind")
        row_index = _nonnegative_int(cell, "rowIndex")
        column_index = _nonnegative_int(cell, "columnIndex")
        effective_row_span = row_span or 1
        effective_column_span = column_span or 1
        if row_index + effective_row_span > row_count or column_index + effective_column_span > column_count:
            raise AzureOCRResponseError("azure_ocr_cell_outside_table_dimensions")
        spans_raw = cell.get("spans", [])
        if not isinstance(spans_raw, list):
            raise AzureOCRResponseError("azure_ocr_cell_spans_must_be_array")
        spans = tuple(_parse_span(_as_mapping(span, "span")) for span in spans_raw)
        for span in spans:
            _validate_span_bounds(span, content_length)
        cells.append(
            OCRTableCell(
                row_index,
                column_index,
                row_span,
                column_span,
                kind,
                _text(cell, "content"),
                spans,
                _parse_regions(cell.get("boundingRegions", []), pages),
            )
        )
    return OCRTable(row_count, column_count, tuple(cells), regions)


def _parse_regions(raw: Any, pages: set[int]) -> tuple[OCRBoundingRegion, ...]:
    if not isinstance(raw, list):
        raise AzureOCRResponseError("azure_ocr_bounding_regions_must_be_array")
    regions: list[OCRBoundingRegion] = []
    for item in raw:
        region = _as_mapping(item, "boundingRegion")
        page_number = _positive_int(region, "pageNumber")
        if page_number not in pages:
            raise AzureOCRResponseError("azure_ocr_bounding_region_unknown_page")
        regions.append(OCRBoundingRegion(page_number, _polygon(region.get("polygon"), "boundingRegion.polygon")))
    return tuple(regions)


def _parse_span(span: Mapping[str, Any]) -> OCRTextSpan:
    return OCRTextSpan(_nonnegative_int(span, "offset"), _nonnegative_int(span, "length"))


def _validate_span_bounds(span: OCRTextSpan, content_length: int) -> None:
    if span.offset + span.length > content_length:
        raise AzureOCRResponseError("azure_ocr_text_span_out_of_bounds")


def _polygon(raw: Any, name: str) -> tuple[float, ...]:
    if not isinstance(raw, list) or len(raw) < 6 or len(raw) % 2:
        raise AzureOCRResponseError(f"azure_ocr_invalid_{name}")
    return tuple(_finite_number(value, name) for value in raw)


def _mapping(value: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    return _as_mapping(value.get(key), key)


def _as_mapping(value: Any, name: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        raise AzureOCRResponseError(f"azure_ocr_{name}_must_be_object")
    return value


def _list(value: Mapping[str, Any], key: str) -> list[Any]:
    result = value.get(key)
    if not isinstance(result, list):
        raise AzureOCRResponseError(f"azure_ocr_{key}_must_be_array")
    return result


def _text(value: Mapping[str, Any], key: str) -> str:
    result = value.get(key)
    if not isinstance(result, str):
        raise AzureOCRResponseError(f"azure_ocr_{key}_must_be_string")
    return result


def _finite_number(value: Any, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise AzureOCRResponseError(f"azure_ocr_{name}_must_be_finite_number")
    return float(value)


def _positive_number(value: Mapping[str, Any], key: str) -> float:
    result = _finite_number(value.get(key), key)
    if result <= 0:
        raise AzureOCRResponseError(f"azure_ocr_{key}_must_be_positive")
    return result


def _nonnegative_int(value: Mapping[str, Any], key: str) -> int:
    result = value.get(key)
    if isinstance(result, bool) or not isinstance(result, int) or result < 0:
        raise AzureOCRResponseError(f"azure_ocr_{key}_must_be_nonnegative_integer")
    return result


def _positive_int(value: Mapping[str, Any], key: str) -> int:
    result = _nonnegative_int(value, key)
    if result == 0:
        raise AzureOCRResponseError(f"azure_ocr_{key}_must_be_positive")
    return result


def _optional_positive_int(value: Mapping[str, Any], key: str) -> int | None:
    return None if key not in value else _positive_int(value, key)
