from __future__ import annotations

from base64 import b64encode
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import ipaddress
import json
import math
import os
import re
import socket
import time
from typing import Any, Mapping
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import HTTPRedirectHandler, Request, build_opener


MISTRAL_OCR_ENDPOINT = "https://api.mistral.ai/v1/ocr"
MISTRAL_OCR_QUALITY_MODEL = "mistral-ocr-latest"
MISTRAL_OCR_TABLE_FORMAT = "html"
MISTRAL_OCR_CONFIDENCE_GRANULARITY = "word"

_BLOCK_TYPES = frozenset(
    {
        "aside_text",
        "caption",
        "code",
        "equation",
        "footer",
        "header",
        "image",
        "list",
        "references",
        "signature",
        "table",
        "text",
        "title",
    }
)
_TRANSIENT_STATUSES = frozenset({408, 409, 425, 429, 500, 502, 503, 504})
_SAFE_CODE = re.compile(r"[A-Za-z0-9_.-]{1,64}\Z")
_ENV_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")
_TABLE_PLACEHOLDER = re.compile(r"\[(?P<id>tbl-\d+\.html)\]\((?P=id)\)")


class MistralOCRError(RuntimeError):
    """Base class for bounded Mistral OCR adapter failures."""


class MistralOCRValidationError(MistralOCRError):
    pass


class MistralOCRConfigError(MistralOCRValidationError):
    pass


class MistralOCRResponseError(MistralOCRValidationError):
    pass


class MistralOCRUnavailableError(MistralOCRError):
    pass


class MistralOCRDeadlineError(MistralOCRUnavailableError):
    pass


class MistralOCRResourceLimitError(MistralOCRError):
    pass


class MistralOCRProviderError(MistralOCRError):
    def __init__(self, status_code: int, provider_code: str, *, retryable: bool):
        self.status_code = status_code
        self.provider_code = provider_code
        self.retryable = retryable
        super().__init__(
            f"mistral_ocr_provider_error:status={status_code}:"
            f"code={provider_code}:retryable={str(retryable).lower()}"
        )


class MistralOCRAuthenticationError(MistralOCRProviderError):
    pass


@dataclass(frozen=True, slots=True)
class MistralOCRConfig:
    endpoint: str
    model: str
    key_env: str
    timeout_seconds: float
    transient_retries: int
    retry_delay_seconds: float
    max_retry_after_seconds: float
    max_input_bytes: int
    max_response_bytes: int

    def __post_init__(self) -> None:
        parsed = urlparse(self.endpoint)
        if not parsed.hostname:
            raise MistralOCRConfigError("mistral_ocr_endpoint_invalid")
        try:
            loopback = ipaddress.ip_address(parsed.hostname).is_loopback
        except ValueError:
            loopback = parsed.hostname.lower() == "localhost"
        if parsed.scheme not in ({"http", "https"} if loopback else {"https"}):
            raise MistralOCRConfigError("mistral_ocr_endpoint_requires_https_or_loopback")
        if parsed.username or parsed.password or parsed.query or parsed.fragment:
            raise MistralOCRConfigError("mistral_ocr_endpoint_invalid")
        try:
            port = parsed.port
        except ValueError:
            raise MistralOCRConfigError("mistral_ocr_endpoint_invalid") from None
        if not loopback and (
            parsed.hostname.lower() != "api.mistral.ai"
            or port not in {None, 443}
            or parsed.path.rstrip("/") != "/v1/ocr"
            or parsed.params
        ):
            raise MistralOCRConfigError("mistral_ocr_endpoint_origin_not_allowed")
        if not isinstance(self.model, str) or not self.model.strip():
            raise MistralOCRConfigError("mistral_ocr_model_required")
        if not isinstance(self.key_env, str) or not _ENV_NAME.fullmatch(self.key_env):
            raise MistralOCRConfigError("mistral_ocr_key_env_invalid")
        _positive_finite(self.timeout_seconds, "timeout")
        if isinstance(self.transient_retries, bool) or not isinstance(self.transient_retries, int) or self.transient_retries < 0:
            raise MistralOCRConfigError("mistral_ocr_retries_must_be_nonnegative")
        _nonnegative_finite(self.retry_delay_seconds, "retry_delay")
        _nonnegative_finite(self.max_retry_after_seconds, "max_retry_after")
        if any(
            isinstance(value, bool) or not isinstance(value, int) or value <= 0
            for value in (self.max_input_bytes, self.max_response_bytes)
        ):
            raise MistralOCRConfigError("mistral_ocr_size_bounds_must_be_positive")


@dataclass(frozen=True, slots=True)
class OCRPageDimensions:
    dpi: int
    pixel_width: int
    pixel_height: int
    pdf_width_points: float
    pdf_height_points: float


@dataclass(frozen=True, slots=True)
class OCRGeometry:
    """Provider geometry plus transparent normalized and PDF-point transforms."""

    provider_bbox: tuple[float, float, float, float] | None
    provider_polygon: tuple[float, ...] | None
    normalized_bbox: tuple[float, float, float, float] | None
    normalized_polygon: tuple[float, ...] | None
    pdf_bbox: tuple[float, float, float, float] | None
    pdf_polygon: tuple[float, ...] | None
    provider_coordinate_space: str = "page_pixels_top_left"
    pdf_coordinate_space: str = "pdf_points_top_left_provider_orientation"
    coordinate_authority: str = "candidate_only"
    requires_independent_anchor: bool = True


@dataclass(frozen=True, slots=True)
class OCRRecognitionEvidence:
    text: str
    confidence: float
    start_index: int
    content_span_text: str
    content_match_kind: str
    geometry: None = None
    evidence_kind: str = "recognition_only"


@dataclass(frozen=True, slots=True)
class OCRAggregateConfidence:
    average: float
    minimum: float


@dataclass(frozen=True, slots=True)
class OCRBlockConfidence:
    average_content: float | None
    minimum_content: float | None
    block_type: float | None


@dataclass(frozen=True, slots=True)
class OCRImageRegion:
    id: str
    geometry: OCRGeometry | None
    annotation: str | None
    content_trust: str = "untrusted_provider_output"


@dataclass(frozen=True, slots=True)
class OCRTable:
    id: str
    content: str
    format: str
    recognition_evidence: tuple[OCRRecognitionEvidence, ...]
    cell_structure_authoritative: bool = False
    content_trust: str = "untrusted_provider_output"


@dataclass(frozen=True, slots=True)
class OCRBlock:
    kind: str
    content: str
    geometry: OCRGeometry | None
    table_id: str | None
    image_id: str | None
    confidence: OCRBlockConfidence | None
    content_trust: str = "untrusted_provider_output"


@dataclass(frozen=True, slots=True)
class OCRPage:
    index: int
    markdown: str
    dimensions: OCRPageDimensions | None
    images: tuple[OCRImageRegion, ...]
    tables: tuple[OCRTable, ...]
    blocks: tuple[OCRBlock, ...]
    recognition_evidence: tuple[OCRRecognitionEvidence, ...]
    confidence: OCRAggregateConfidence | None
    content_trust: str = "untrusted_provider_output"


@dataclass(frozen=True, slots=True)
class OCRUsage:
    pages_processed: int
    document_size_bytes: int | None


@dataclass(frozen=True, slots=True)
class OCRResult:
    provider: str
    requested_model: str
    provider_model: str
    pages: tuple[OCRPage, ...]
    usage: OCRUsage


class _RejectRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):  # type: ignore[no-untyped-def]
        return None


class MistralDocumentOCR:
    """Optional OCR adapter; credentials are resolved only from a named env var."""

    def __init__(self, config: MistralOCRConfig):
        self.config = config

    def analyze(self, document_bytes: bytes, *, deadline: float | None = None) -> OCRResult:
        if not isinstance(document_bytes, bytes):
            raise TypeError("document_bytes must be bytes")
        if not document_bytes:
            raise MistralOCRResourceLimitError("mistral_ocr_empty_input")
        if len(document_bytes) > self.config.max_input_bytes:
            raise MistralOCRResourceLimitError("mistral_ocr_input_size_exceeded")
        api_key = os.environ.get(self.config.key_env)
        if not api_key:
            raise MistralOCRAuthenticationError(0, "key_env_unavailable", retryable=False)

        local_deadline = time.monotonic() + self.config.timeout_seconds
        if deadline is not None:
            if isinstance(deadline, bool) or not isinstance(deadline, (int, float)) or not math.isfinite(deadline):
                raise MistralOCRValidationError("mistral_ocr_deadline_invalid")
            local_deadline = min(local_deadline, deadline)
        if local_deadline <= time.monotonic():
            raise MistralOCRDeadlineError("mistral_ocr_deadline_exceeded")

        encoded = b64encode(document_bytes).decode("ascii")
        payload = {
            "model": self.config.model,
            "document": {
                "type": "document_url",
                "document_url": f"data:application/pdf;base64,{encoded}",
            },
            "table_format": MISTRAL_OCR_TABLE_FORMAT,
            "include_image_base64": False,
            "include_blocks": True,
            "confidence_scores_granularity": MISTRAL_OCR_CONFIDENCE_GRANULARITY,
        }
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        request = Request(
            self.config.endpoint,
            data=body,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            method="POST",
        )
        response_body = self._request(request, deadline=local_deadline)
        parsed = _decode_json(response_body)
        return _parse_result(parsed, requested_model=self.config.model)

    def _request(self, request: Request, *, deadline: float) -> bytes:
        for attempt in range(self.config.transient_retries + 1):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise MistralOCRDeadlineError("mistral_ocr_deadline_exceeded")
            try:
                opener = build_opener(_RejectRedirects())
                with opener.open(request, timeout=remaining) as response:
                    status = response.getcode()
                    if status != 200:
                        raise MistralOCRProviderError(status, "unexpected_status", retryable=False)
                    content_type = response.headers.get("Content-Type", "").partition(";")[0].strip().lower()
                    if content_type != "application/json":
                        raise MistralOCRResponseError("mistral_ocr_response_not_json")
                    return _read_bounded(response, self.config.max_response_bytes)
            except HTTPError as exc:
                error_body = _read_bounded(exc, self.config.max_response_bytes)
                provider_code = _provider_code(error_body)
                retryable = exc.code in _TRANSIENT_STATUSES
                error_type = MistralOCRAuthenticationError if exc.code in {401, 403} else MistralOCRProviderError
                error = error_type(exc.code, provider_code, retryable=retryable)
                if not retryable or attempt == self.config.transient_retries:
                    raise error from exc
                self._wait_before_retry(deadline, exc.headers.get("Retry-After"), error)
            except (URLError, TimeoutError, socket.timeout) as exc:
                if attempt == self.config.transient_retries:
                    raise MistralOCRUnavailableError(
                        f"mistral_ocr_transport_unavailable:{type(exc).__name__}"
                    ) from exc
                self._wait_before_retry(deadline, None, None)
        raise MistralOCRUnavailableError("mistral_ocr_request_failed:no_response")

    def _wait_before_retry(
        self,
        deadline: float,
        retry_after: str | None,
        provider_error: MistralOCRProviderError | None,
    ) -> None:
        delay = self.config.retry_delay_seconds
        if retry_after is not None:
            parsed = _parse_retry_after(retry_after)
            if parsed is not None:
                if parsed > self.config.max_retry_after_seconds:
                    if provider_error is not None:
                        raise provider_error
                    raise MistralOCRUnavailableError("mistral_ocr_retry_after_exceeds_bound")
                delay = parsed
        remaining = deadline - time.monotonic()
        if remaining <= 0 or delay >= remaining:
            raise MistralOCRDeadlineError("mistral_ocr_deadline_exceeded")
        time.sleep(delay)


def _read_bounded(response: Any, maximum: int) -> bytes:
    length = response.headers.get("Content-Length") if getattr(response, "headers", None) is not None else None
    if length is not None:
        try:
            if int(length) > maximum:
                raise MistralOCRResourceLimitError("mistral_ocr_response_size_exceeded")
        except ValueError:
            raise MistralOCRResponseError("mistral_ocr_invalid_content_length") from None
    body = response.read(maximum + 1)
    if len(body) > maximum:
        raise MistralOCRResourceLimitError("mistral_ocr_response_size_exceeded")
    return body


def _decode_json(body: bytes) -> dict[str, Any]:
    def object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise MistralOCRResponseError("mistral_ocr_duplicate_json_key")
            result[key] = value
        return result

    try:
        value = json.loads(
            body.decode("utf-8"),
            object_pairs_hook=object_pairs,
            parse_constant=lambda _: (_ for _ in ()).throw(ValueError("nonfinite")),
        )
    except MistralOCRResponseError:
        raise
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError, RecursionError) as exc:
        raise MistralOCRResponseError("mistral_ocr_invalid_json") from exc
    if not isinstance(value, dict):
        raise MistralOCRResponseError("mistral_ocr_response_must_be_object")
    return value


def _provider_code(body: bytes) -> str:
    try:
        parsed = _decode_json(body)
    except MistralOCRValidationError:
        return "unknown"
    error = parsed.get("error")
    code: object = error.get("code") if isinstance(error, dict) else parsed.get("code")
    return code if isinstance(code, str) and _SAFE_CODE.fullmatch(code) else "unknown"


def _parse_retry_after(value: str) -> float | None:
    try:
        delay = float(value)
    except ValueError:
        try:
            when = parsedate_to_datetime(value)
            if when.tzinfo is None:
                when = when.replace(tzinfo=timezone.utc)
            delay = max(0.0, (when - datetime.now(timezone.utc)).total_seconds())
        except (TypeError, ValueError, OverflowError):
            return None
    return delay if math.isfinite(delay) and delay >= 0 else None


def _parse_result(payload: Mapping[str, Any], *, requested_model: str) -> OCRResult:
    provider_model = _text(payload, "model")
    if not provider_model.strip():
        raise MistralOCRResponseError("mistral_ocr_provider_model_empty")
    pages_raw = _array(payload, "pages")
    page_indices: set[int] = set()
    pages: list[OCRPage] = []
    for raw in pages_raw:
        page = _as_mapping(raw, "page")
        index = _nonnegative_int(page, "index")
        if index in page_indices:
            raise MistralOCRResponseError("mistral_ocr_duplicate_page_index")
        page_indices.add(index)
        pages.append(_parse_page(page, index=index))

    usage_raw = _mapping(payload, "usage_info")
    pages_processed = _nonnegative_int(usage_raw, "pages_processed")
    if pages_processed != len(pages):
        raise MistralOCRResponseError("mistral_ocr_usage_page_count_mismatch")
    document_size = usage_raw.get("doc_size_bytes")
    if document_size is not None:
        document_size = _nonnegative_int(usage_raw, "doc_size_bytes")
    return OCRResult(
        provider="mistral_document_ai",
        requested_model=requested_model,
        provider_model=provider_model,
        pages=tuple(pages),
        usage=OCRUsage(pages_processed, document_size),
    )


def _parse_page(page: Mapping[str, Any], *, index: int) -> OCRPage:
    markdown = _text(page, "markdown")
    dimensions = _parse_dimensions(page.get("dimensions"))

    images: list[OCRImageRegion] = []
    image_ids: set[str] = set()
    for raw in _array(page, "images"):
        image = _as_mapping(raw, "image")
        image_id = _text(image, "id")
        if not image_id or image_id in image_ids:
            raise MistralOCRResponseError("mistral_ocr_duplicate_or_empty_image_id")
        image_ids.add(image_id)
        if image.get("image_base64") not in {None, ""}:
            raise MistralOCRResponseError("mistral_ocr_unexpected_image_base64")
        annotation = image.get("image_annotation")
        if annotation is not None and not isinstance(annotation, str):
            raise MistralOCRResponseError("mistral_ocr_image_annotation_must_be_string")
        images.append(OCRImageRegion(image_id, _parse_geometry(image, dimensions), annotation))

    tables_raw = page.get("tables", [])
    if not isinstance(tables_raw, list):
        raise MistralOCRResponseError("mistral_ocr_tables_must_be_array")
    tables: list[OCRTable] = []
    table_ids: set[str] = set()
    for raw in tables_raw:
        table = _as_mapping(raw, "table")
        table_id = _text(table, "id")
        if not table_id or table_id in table_ids:
            raise MistralOCRResponseError("mistral_ocr_duplicate_or_empty_table_id")
        table_ids.add(table_id)
        content = _text(table, "content")
        table_format = _text(table, "format")
        if table_format != MISTRAL_OCR_TABLE_FORMAT:
            raise MistralOCRResponseError("mistral_ocr_unexpected_table_format")
        evidence = _parse_recognition_evidence(table.get("word_confidence_scores"), content)
        tables.append(OCRTable(table_id, content, table_format, evidence))

    blocks_raw = page.get("blocks")
    if blocks_raw is None:
        blocks_raw = []
    if not isinstance(blocks_raw, list):
        raise MistralOCRResponseError("mistral_ocr_blocks_must_be_array")
    blocks: list[OCRBlock] = []
    referenced_tables: set[str] = set()
    referenced_images: set[str] = set()
    for raw in blocks_raw:
        block = _as_mapping(raw, "block")
        kind = _text(block, "type")
        if kind not in _BLOCK_TYPES:
            raise MistralOCRResponseError("mistral_ocr_unknown_block_type")
        table_id = _optional_text(block, "table_id")
        image_id = _optional_text(block, "image_id")
        if kind == "table" and table_id is not None:
            if table_id not in table_ids or table_id in referenced_tables:
                raise MistralOCRResponseError("mistral_ocr_invalid_or_duplicate_table_reference")
            referenced_tables.add(table_id)
        elif table_id is not None:
            raise MistralOCRResponseError("mistral_ocr_table_id_on_non_table_block")
        if kind == "image" and image_id is not None:
            if image_id not in image_ids or image_id in referenced_images:
                raise MistralOCRResponseError("mistral_ocr_invalid_or_duplicate_image_reference")
            referenced_images.add(image_id)
        elif image_id is not None:
            raise MistralOCRResponseError("mistral_ocr_image_id_on_non_image_block")
        blocks.append(
            OCRBlock(
                kind=kind,
                content=_text(block, "content"),
                geometry=_parse_geometry(block, dimensions),
                table_id=table_id,
                image_id=image_id,
                confidence=_parse_block_confidence(block.get("confidence_scores")),
            )
        )

    page_confidence, evidence = _parse_page_confidence(
        page.get("confidence_scores"), markdown, table_placeholder_ids=frozenset(table_ids)
    )
    return OCRPage(
        index=index,
        markdown=markdown,
        dimensions=dimensions,
        images=tuple(images),
        tables=tuple(tables),
        blocks=tuple(blocks),
        recognition_evidence=evidence,
        confidence=page_confidence,
    )


def _parse_dimensions(raw: Any) -> OCRPageDimensions | None:
    if raw is None:
        return None
    data = _as_mapping(raw, "dimensions")
    dpi = _positive_int(data, "dpi")
    width = _positive_int(data, "width")
    height = _positive_int(data, "height")
    scale = 72.0 / dpi
    return OCRPageDimensions(dpi, width, height, width * scale, height * scale)


def _parse_geometry(raw: Mapping[str, Any], dimensions: OCRPageDimensions | None) -> OCRGeometry | None:
    coordinate_keys = ("top_left_x", "top_left_y", "bottom_right_x", "bottom_right_y")
    supplied = tuple(key in raw and raw.get(key) is not None for key in coordinate_keys)
    bbox_raw = raw.get("bbox")
    polygon_raw = raw.get("polygon")
    if any(supplied) and not all(supplied):
        raise MistralOCRResponseError("mistral_ocr_partial_bbox")
    if all(supplied) and bbox_raw is not None:
        raise MistralOCRResponseError("mistral_ocr_duplicate_bbox_representation")
    bbox: tuple[float, float, float, float] | None
    if all(supplied):
        bbox = tuple(_finite_number(raw[key], key) for key in coordinate_keys)  # type: ignore[assignment]
    elif bbox_raw is not None:
        bbox_values = _number_sequence(bbox_raw, "bbox", exact_length=4)
        bbox = (bbox_values[0], bbox_values[1], bbox_values[2], bbox_values[3])
    else:
        bbox = None
    polygon = _number_sequence(polygon_raw, "polygon", minimum_length=6, even=True) if polygon_raw is not None else None
    if bbox is None and polygon is not None:
        points = tuple(zip(polygon[0::2], polygon[1::2]))
        bbox = (
            min(point[0] for point in points),
            min(point[1] for point in points),
            max(point[0] for point in points),
            max(point[1] for point in points),
        )
    if bbox is None and polygon is None:
        return None
    assert bbox is not None
    if bbox[0] < 0 or bbox[1] < 0 or bbox[2] < bbox[0] or bbox[3] < bbox[1]:
        raise MistralOCRResponseError("mistral_ocr_geometry_out_of_bounds")
    if polygon is not None and any(value < 0 for value in polygon):
        raise MistralOCRResponseError("mistral_ocr_geometry_out_of_bounds")

    normalized_bbox = None
    normalized_polygon = None
    pdf_bbox = None
    pdf_polygon = None
    if dimensions is not None:
        if bbox[2] > dimensions.pixel_width or bbox[3] > dimensions.pixel_height:
            raise MistralOCRResponseError("mistral_ocr_geometry_out_of_bounds")
        if polygon is not None and any(
            value > (dimensions.pixel_width if offset % 2 == 0 else dimensions.pixel_height)
            for offset, value in enumerate(polygon)
        ):
            raise MistralOCRResponseError("mistral_ocr_geometry_out_of_bounds")
        normalized_bbox = (
            bbox[0] / dimensions.pixel_width,
            bbox[1] / dimensions.pixel_height,
            bbox[2] / dimensions.pixel_width,
            bbox[3] / dimensions.pixel_height,
        )
        if polygon is not None:
            normalized_polygon = tuple(
                value / (dimensions.pixel_width if offset % 2 == 0 else dimensions.pixel_height)
                for offset, value in enumerate(polygon)
            )
        scale = 72.0 / dimensions.dpi
        pdf_bbox = tuple(value * scale for value in bbox)  # type: ignore[assignment]
        if polygon is not None:
            pdf_polygon = tuple(value * scale for value in polygon)
    return OCRGeometry(bbox, polygon, normalized_bbox, normalized_polygon, pdf_bbox, pdf_polygon)


def _parse_page_confidence(
    raw: Any,
    content: str,
    *,
    table_placeholder_ids: frozenset[str],
) -> tuple[OCRAggregateConfidence | None, tuple[OCRRecognitionEvidence, ...]]:
    if raw is None:
        return None, ()
    data = _as_mapping(raw, "page_confidence_scores")
    average = _confidence(data.get("average_page_confidence_score"), "average_page_confidence_score")
    minimum = _confidence(data.get("minimum_page_confidence_score"), "minimum_page_confidence_score")
    if minimum > average:
        raise MistralOCRResponseError("mistral_ocr_page_confidence_order_invalid")
    return OCRAggregateConfidence(average, minimum), _parse_recognition_evidence(
        data.get("word_confidence_scores"),
        content,
        table_placeholder_ids=table_placeholder_ids,
    )


def _parse_block_confidence(raw: Any) -> OCRBlockConfidence | None:
    if raw is None:
        return None
    data = _as_mapping(raw, "block_confidence_scores")
    average = _nullable_confidence(data.get("average_content_confidence_score"), "average_content")
    minimum = _nullable_confidence(data.get("minimum_content_confidence_score"), "minimum_content")
    block_type = _nullable_confidence(data.get("block_type_confidence_score"), "block_type")
    if average is not None and minimum is not None and minimum > average:
        raise MistralOCRResponseError("mistral_ocr_block_confidence_order_invalid")
    return OCRBlockConfidence(average, minimum, block_type)


def _parse_recognition_evidence(
    raw: Any,
    content: str,
    *,
    table_placeholder_ids: frozenset[str] = frozenset(),
) -> tuple[OCRRecognitionEvidence, ...]:
    if raw is None:
        return ()
    if not isinstance(raw, list):
        raise MistralOCRResponseError("mistral_ocr_word_confidence_scores_must_be_array")
    evidence: list[OCRRecognitionEvidence] = []
    seen: set[tuple[int, str]] = set()
    previous_end = 0
    for item in raw:
        data = _as_mapping(item, "word_confidence")
        text = _text(data, "text")
        if not text:
            raise MistralOCRResponseError("mistral_ocr_word_confidence_text_empty")
        start = _nonnegative_int(data, "start_index")
        if start > len(content):
            raise MistralOCRResponseError("mistral_ocr_word_confidence_out_of_bounds")
        identity = (start, text)
        if identity in seen:
            raise MistralOCRResponseError("mistral_ocr_duplicate_word_confidence")
        seen.add(identity)
        content_span_text, match_kind = _resolve_confidence_span(
            content,
            start,
            text,
            table_placeholder_ids=table_placeholder_ids,
        )
        if start < previous_end:
            raise MistralOCRResponseError("mistral_ocr_word_confidence_order_invalid")
        previous_end = start + len(content_span_text)
        evidence.append(
            OCRRecognitionEvidence(
                text,
                _confidence(data.get("confidence"), "word"),
                start,
                content_span_text,
                match_kind,
            )
        )
    return tuple(evidence)


def _resolve_confidence_span(
    content: str,
    start: int,
    reported_text: str,
    *,
    table_placeholder_ids: frozenset[str],
) -> tuple[str, str]:
    if content.startswith(reported_text, start):
        return reported_text, "exact"
    reported_placeholder = _TABLE_PLACEHOLDER.fullmatch(reported_text)
    content_placeholder = _TABLE_PLACEHOLDER.match(content, start)
    if (
        reported_placeholder is not None
        and content_placeholder is not None
        and content_placeholder.group("id") in table_placeholder_ids
    ):
        return content_placeholder.group(0), "rebased_table_placeholder"
    raise MistralOCRResponseError("mistral_ocr_word_confidence_text_mismatch")


def _mapping(data: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    return _as_mapping(data.get(key), key)


def _as_mapping(value: Any, name: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        raise MistralOCRResponseError(f"mistral_ocr_{name}_must_be_object")
    return value


def _array(data: Mapping[str, Any], key: str) -> list[Any]:
    value = data.get(key)
    if not isinstance(value, list):
        raise MistralOCRResponseError(f"mistral_ocr_{key}_must_be_array")
    return value


def _text(data: Mapping[str, Any], key: str) -> str:
    value = data.get(key)
    if not isinstance(value, str):
        raise MistralOCRResponseError(f"mistral_ocr_{key}_must_be_string")
    return value


def _optional_text(data: Mapping[str, Any], key: str) -> str | None:
    value = data.get(key)
    if value is None:
        return None
    if not isinstance(value, str) or not value:
        raise MistralOCRResponseError(f"mistral_ocr_{key}_must_be_nonempty_string")
    return value


def _finite_number(value: Any, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise MistralOCRResponseError(f"mistral_ocr_{name}_must_be_finite_number")
    return float(value)


def _positive_finite(value: Any, name: str) -> None:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise MistralOCRConfigError(f"mistral_ocr_{name}_must_be_positive")


def _nonnegative_finite(value: Any, name: str) -> None:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0:
        raise MistralOCRConfigError(f"mistral_ocr_{name}_must_be_nonnegative")


def _positive_int(data: Mapping[str, Any], key: str) -> int:
    value = data.get(key)
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise MistralOCRResponseError(f"mistral_ocr_{key}_must_be_positive_integer")
    return value


def _nonnegative_int(data: Mapping[str, Any], key: str) -> int:
    value = data.get(key)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise MistralOCRResponseError(f"mistral_ocr_{key}_must_be_nonnegative_integer")
    return value


def _confidence(value: Any, name: str) -> float:
    number = _finite_number(value, name)
    if number < 0 or number > 1:
        raise MistralOCRResponseError(f"mistral_ocr_{name}_confidence_out_of_range")
    return number


def _nullable_confidence(value: Any, name: str) -> float | None:
    return None if value is None else _confidence(value, name)


def _number_sequence(
    value: Any,
    name: str,
    *,
    exact_length: int | None = None,
    minimum_length: int | None = None,
    even: bool = False,
) -> tuple[float, ...]:
    if not isinstance(value, list):
        raise MistralOCRResponseError(f"mistral_ocr_{name}_must_be_array")
    if exact_length is not None and len(value) != exact_length:
        raise MistralOCRResponseError(f"mistral_ocr_{name}_invalid_length")
    if minimum_length is not None and len(value) < minimum_length:
        raise MistralOCRResponseError(f"mistral_ocr_{name}_invalid_length")
    if even and len(value) % 2:
        raise MistralOCRResponseError(f"mistral_ocr_{name}_invalid_length")
    return tuple(_finite_number(item, name) for item in value)
