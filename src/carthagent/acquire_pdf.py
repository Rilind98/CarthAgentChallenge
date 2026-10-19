from __future__ import annotations

import math
from pathlib import Path
import re
from difflib import SequenceMatcher
import shutil
import subprocess
import tempfile
import time
from typing import Callable, Protocol

import pymupdf as fitz

from .acquire_xlsx import AcquisitionResult, ResourceLimitError
from .geometry import PageTransform
from .models import CoverageDisposition, CoverageItem, SourceObservation
from .numeric import NumericParseError, parse_amount
from .ocr_azure import AzureDocumentIntelligenceOCR, AzureOCRError, OCRPage, OCRTable, OCRTableCell, OCRWord
from .ocr_mistral import MistralDocumentOCR, MistralOCRError, OCRPage as MistralOCRPage
from .policies import ResourceBounds
from .structure import FIELDS, HEADER_GROUPS, IDENTITIES, match, norm


class OCRProvider(Protocol):
    def acquire(self, page: fitz.Page, document_id: str, source_hash: str, page_index: int, bounds: ResourceBounds, deadline: float) -> tuple[list[SourceObservation], list[str]]: ...


class LocalTesseractOCR:
    def acquire(self, page: fitz.Page, document_id: str, source_hash: str, page_index: int, bounds: ResourceBounds, deadline: float) -> tuple[list[SourceObservation], list[str]]:
        return _ocr_page(page, document_id, source_hash, page_index, bounds, deadline)


OCR_PROVIDERS: dict[str, OCRProvider] = {"tesseract": LocalTesseractOCR()}


def _same_visual_baseline(group_center: float, group_height: float, bbox: tuple[float, float, float, float]) -> bool:
    """Return true only for glyphs on the same printed baseline."""
    center = (bbox[1] + bbox[3]) / 2
    height = max(bbox[3] - bbox[1], 0.01)
    return abs(center - group_center) <= min(group_height, height) * 0.35


def _semantic_spans(document_id: str, source_hash: str, page_index: int, row_id: str, row_words: list[SourceObservation], channel: str) -> list[SourceObservation]:
    generated: list[SourceObservation] = []
    for role, vocabulary in (("semantic_field", FIELDS), ("semantic_identity", IDENTITIES), ("semantic_group", HEADER_GROUPS)):
        spans: dict[str, tuple[int, int]] = {}
        for start in range(len(row_words)):
            for end in range(start + 1, len(row_words) + 1):
                value = match(" ".join(item.text or "" for item in row_words[start:end]), vocabulary)
                if value is not None and (value not in spans or end - start < spans[value][1] - spans[value][0]):
                    spans[value] = (start, end)
        generic = spans.get("amount")
        if role == "semantic_field" and generic and any(value != "amount" and max(generic[0], span[0]) < min(generic[1], span[1]) for value, span in spans.items()):
            spans.pop("amount")
        for value, (start, end) in spans.items():
            span = row_words[start:end]
            boxes = [item.bbox for item in span if item.bbox]
            bbox = (min(box[0] for box in boxes), min(box[1] for box in boxes), max(box[2] for box in boxes), max(box[3] for box in boxes))
            generated.append(SourceObservation(
                f"{row_id}:header:{role}:{value}:{start}", document_id, source_hash, "logical_row",
                f"page:{page_index + 1}:header:{role}:{value}:{start}", channel,
                " ".join(item.text or "" for item in span), bbox,
                {"member_ids": [item.id for item in span], "page_index": page_index, "structure_role": "native_line", role: value},
            ))
    return generated


def register_ocr_provider(name: str, provider: OCRProvider) -> None:
    if not name or name in OCR_PROVIDERS:
        raise ValueError("OCR provider name must be nonempty and unused")
    OCR_PROVIDERS[name] = provider


def _ocr_page(page: fitz.Page, document_id: str, source_hash: str, page_index: int, bounds: ResourceBounds, deadline: float) -> tuple[list[SourceObservation], list[str]]:
    if not shutil.which("tesseract"):
        return [], [f"page:{page_index + 1}:ocr_unavailable"]
    zoom = bounds.ocr_dpi / 72.0
    canonical = page.cropbox
    pixels = int(canonical.width * zoom) * int(canonical.height * zoom)
    if pixels > bounds.max_page_megapixels * 1_000_000:
        return [], [f"page:{page_index + 1}:render_pixel_limit"]
    # Render in the page's canonical, unrotated coordinate space. This keeps
    # OCR upright for 180/270 metadata rotations and makes returned boxes
    # directly comparable with native PyMuPDF word coordinates.
    transform = PageTransform(scale=zoom, rotation=0, page_width=canonical.width, page_height=canonical.height)
    with tempfile.TemporaryDirectory(prefix="carthagent-ocr-") as tmp:
        image_path = Path(tmp) / "page.png"
        page.get_pixmap(matrix=fitz.Matrix(zoom, zoom).prerotate(-page.rotation), alpha=False).save(image_path)
        proc = None
        for attempt in range(bounds.transient_retries + 1):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return [], [f"page:{page_index + 1}:ocr_document_timeout"]
            try:
                proc = subprocess.run(["tesseract", str(image_path), "stdout", "-l", "deu+eng", "tsv"], capture_output=True, text=True, timeout=remaining)
            except (subprocess.TimeoutExpired, OSError) as exc:
                if attempt == bounds.transient_retries:
                    return [], [f"page:{page_index + 1}:ocr_failed:{type(exc).__name__}"]
                continue
            if proc.returncode == 0:
                break
            if attempt == bounds.transient_retries:
                return [], [f"page:{page_index + 1}:ocr_failed:{proc.returncode}"]
        assert proc is not None
        observations: list[SourceObservation] = []
        line_groups: dict[tuple[str, str, str], list[SourceObservation]] = {}
        for line_no, line in enumerate(proc.stdout.splitlines()[1:], 2):
            parts = line.split("\t", 11)
            if len(parts) != 12 or not parts[11].strip():
                continue
            left, top, width, height = map(float, parts[6:10])
            rendered = (left, top, left + width, top + height)
            source_box = transform.to_source(rendered)
            oid = f"{document_id}:pdf:p{page_index + 1}:ocr:{line_no}"
            observation = SourceObservation(oid, document_id, source_hash, "word", f"page:{page_index + 1}:ocr:{line_no}", "pdf_ocr", parts[11], source_box, {"confidence": parts[10], "render_box": rendered, "transform": {"scale": zoom, "rotation": page.rotation, "deskew": 0.0}, "roundtrip_ok": transform.roundtrip_ok(source_box), "page_index": page_index, "page_rotation": page.rotation, "ocr_block": parts[2], "ocr_paragraph": parts[3], "ocr_line": parts[4]})
            observations.append(observation)
            line_groups.setdefault((parts[2], parts[3], parts[4]), []).append(observation)
            if len(observations) > bounds.max_cells:
                raise ResourceLimitError(f"ocr_observation_limit:{len(observations)}>{bounds.max_cells}")
        for row_no, row_words in enumerate(line_groups.values()):
            row_words.sort(key=lambda item: item.bbox[0] if item.bbox else 0)
            boxes = [item.bbox for item in row_words if item.bbox]
            bbox = (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))
            observations.append(SourceObservation(f"{document_id}:pdf:p{page_index + 1}:ocrrow:{row_no}", document_id, source_hash, "logical_row", f"page:{page_index + 1}:ocrrow:{row_no}", "pdf_ocr_structure", " ".join(item.text or "" for item in row_words), bbox, {"member_ids": [item.id for item in row_words], "page_index": page_index, "page_rotation": page.rotation, "structure_role": "visual_row"}))
            observations.extend(item for item in _semantic_spans(document_id, source_hash, page_index, f"{document_id}:pdf:p{page_index + 1}:ocrrow:{row_no}", row_words, "pdf_ocr_structure") if item.metadata.get("semantic_group"))
            # Enumerate contiguous header spans and retain the narrowest span
            # for each known field. This reconstructs multi-token headers while
            # keeping each amount column separate from the full visual row.
            field_spans: dict[str, tuple[int, int]] = {}
            for start in range(len(row_words)):
                for end in range(start + 1, len(row_words) + 1):
                    field = match(" ".join(item.text or "" for item in row_words[start:end]), FIELDS)
                    if field is not None and (field not in field_spans or end - start < field_spans[field][1] - field_spans[field][0]):
                        field_spans[field] = (start, end)
            generic = field_spans.get("amount")
            if generic and any(field != "amount" and max(generic[0], span[0]) < min(generic[1], span[1]) for field, span in field_spans.items()):
                field_spans.pop("amount")
            for field, (start, end) in field_spans.items():
                span = row_words[start:end]
                span_boxes = [item.bbox for item in span if item.bbox]
                span_bbox = (min(b[0] for b in span_boxes), min(b[1] for b in span_boxes), max(b[2] for b in span_boxes), max(b[3] for b in span_boxes))
                observations.append(SourceObservation(f"{document_id}:pdf:p{page_index + 1}:ocrheader:{row_no}:{field}", document_id, source_hash, "logical_row", f"page:{page_index + 1}:ocrheader:{row_no}:{field}", "pdf_ocr_structure", " ".join(item.text or "" for item in span), span_bbox, {"member_ids": [item.id for item in span], "page_index": page_index, "page_rotation": page.rotation, "structure_role": "native_line", "semantic_field": field}))
            identity_spans: dict[str, tuple[int, int]] = {}
            for start in range(len(row_words)):
                for end in range(start + 1, len(row_words) + 1):
                    axis = match(" ".join(item.text or "" for item in row_words[start:end]), IDENTITIES)
                    if axis is not None and (axis not in identity_spans or end - start < identity_spans[axis][1] - identity_spans[axis][0]):
                        identity_spans[axis] = (start, end)
            for axis, (start, end) in identity_spans.items():
                span = row_words[start:end]
                span_boxes = [item.bbox for item in span if item.bbox]
                span_bbox = (min(b[0] for b in span_boxes), min(b[1] for b in span_boxes), max(b[2] for b in span_boxes), max(b[3] for b in span_boxes))
                observations.append(SourceObservation(f"{document_id}:pdf:p{page_index + 1}:ocridentity:{row_no}:{axis}", document_id, source_hash, "logical_row", f"page:{page_index + 1}:ocridentity:{row_no}:{axis}", "pdf_ocr_structure", " ".join(item.text or "" for item in span), span_bbox, {"member_ids": [item.id for item in span], "page_index": page_index, "page_rotation": page.rotation, "structure_role": "native_line", "semantic_identity": axis}))
        return observations, []


def acquire_pdf(path: Path, document_id: str, source_hash: str, bounds: ResourceBounds, *, ocr_provider: str = "tesseract", ocr_enabled: bool = True, deadline: float | None = None, azure_provider: AzureDocumentIntelligenceOCR | None = None, mistral_provider: MistralDocumentOCR | None = None) -> AcquisitionResult:
    deadline = deadline if deadline is not None else time.monotonic() + bounds.document_timeout_seconds
    try:
        doc = fitz.open(path)
    except Exception as exc:
        raise ValueError(f"corrupt_pdf:{type(exc).__name__}") from exc
    if doc.page_count > bounds.max_pdf_pages:
        raise ResourceLimitError(f"page_limit:{doc.page_count}>{bounds.max_pdf_pages}")
    observations: list[SourceObservation] = []
    coverage: list[CoverageItem] = []
    diagnostics: list[str] = []
    partial = False
    azure_pages: dict[int, list[SourceObservation]] = {}
    azure_failed = False
    mistral_pages: dict[int, MistralOCRPage] = {}
    if ocr_enabled and ocr_provider == "azure_document_intelligence":
        # Azure is the primary source for scan pages (no text layer, or a
        # raster with almost no text) and outline pages (vector glyphs, no
        # text layer). Only those pages are sent; native pages keep the
        # native path. Tesseract is the fallback when Azure cannot answer.
        cloud_pages = [index for index, page in enumerate(doc) if not page.get_text("words") or bool(page.get_images(full=True) and len(page.get_text("words")) < 10)]
        if cloud_pages and azure_provider is None:
            diagnostics.append("azure_ocr_provider_unconfigured"); diagnostics.append("azure_unavailable_fallback_tesseract:unconfigured"); azure_failed = True; partial = True
        elif cloud_pages:
            # Polling is bounded by what is left of the document budget after
            # the reserve the later phases need.
            azure_deadline = deadline - bounds.post_mapper_reserve_seconds
            if azure_deadline <= time.monotonic():
                diagnostics.append("azure_unavailable_fallback_tesseract:document_budget_exhausted"); azure_failed = True; partial = True
            else:
                try:
                    # Upload only the scan/outline pages (request-size limits,
                    # billing); uploaded page k is original page cloud_pages[k - 1].
                    with fitz.open() as subset:
                        for index in cloud_pages:
                            subset.insert_pdf(doc, from_page=index, to_page=index)
                        subset_bytes = subset.tobytes()
                    result = azure_provider.analyze(subset_bytes, deadline=azure_deadline)
                    for provider_page in result.pages:
                        position = provider_page.page_number - 1
                        if 0 <= position < len(cloud_pages):
                            page_index = cloud_pages[position]
                            azure_pages[page_index], page_issues = _azure_page_observations(provider_page, doc[page_index], document_id, source_hash, page_index, result.tables)
                            diagnostics.extend(page_issues)
                    # A requested page Azure did not return (the free tier, for
                    # one, analyses only the first two pages) goes to Tesseract.
                    diagnostics.extend(f"page:{index + 1}:azure_page_missing_fallback_tesseract" for index in cloud_pages if not azure_pages.get(index))
                except (AzureOCRError, OSError, ValueError) as exc:
                    diagnostics.append(f"azure_unavailable_fallback_tesseract:{type(exc).__name__}:{exc}"); azure_failed = True; partial = True
    if ocr_enabled and ocr_provider == "mistral_document_ai":
        if mistral_provider is None:
            diagnostics.append("mistral_ocr_provider_unconfigured"); partial = True
        else:
            try:
                result = mistral_provider.analyze(path.read_bytes(), deadline=deadline)
                mistral_pages = {page.index: page for page in result.pages}
            except (MistralOCRError, OSError, ValueError) as exc:
                diagnostics.append(f"mistral_ocr_unavailable:{type(exc).__name__}:{exc}"); partial = True
    for i, page in enumerate(doc):
        if time.monotonic() >= deadline:
            diagnostics.append(f"document_timeout:page:{i + 1}")
            for remaining_page in range(i, doc.page_count):
                coverage.append(CoverageItem(
                    f"{document_id}:coverage:page:{remaining_page + 1}", document_id,
                    "pdf_page", f"page:{remaining_page + 1}", CoverageDisposition.UNRESOLVED,
                    reason="unprocessed_document_timeout",
                ))
            partial = True
            break
        page_id = f"{document_id}:coverage:page:{i + 1}"
        words = page.get_text("words", sort=True)
        drawings = page.get_drawings()
        images = page.get_images(full=True)
        line_members: dict[tuple[int, int], list[str]] = {}
        line_words: dict[tuple[int, int], list[str]] = {}
        line_boxes: dict[tuple[int, int], list[tuple[float, float, float, float]]] = {}
        for word_no, word in enumerate(words):
            x0, y0, x1, y1, text = word[:5]
            oid = f"{document_id}:pdf:p{i + 1}:native:{word_no}"
            observations.append(SourceObservation(oid, document_id, source_hash, "word", f"page:{i + 1}:word:{word_no}", "pdf_native", text, (x0, y0, x1, y1), {"page_index": i, "page_rotation": page.rotation, "page_width": page.rect.width, "page_height": page.rect.height}))
            coverage.append(CoverageItem(f"{document_id}:coverage:word:p{i + 1}:{word_no}", document_id, "pdf_word", f"page:{i + 1}:word:{word_no}", CoverageDisposition.UNRESOLVED, reason="awaiting_semantic_disposition"))
            if len(observations) > bounds.max_cells:
                raise ResourceLimitError(f"pdf_observation_limit:{len(observations)}>{bounds.max_cells}")
            line_key = (int(word[5]), int(word[6]))
            line_members.setdefault(line_key, []).append(oid); line_words.setdefault(line_key, []).append(text); line_boxes.setdefault(line_key, []).append((x0, y0, x1, y1))
        for (block_no, line_no), members in line_members.items():
            boxes = line_boxes[(block_no, line_no)]
            bbox = (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))
            observations.append(SourceObservation(f"{document_id}:pdf:p{i + 1}:row:{block_no}:{line_no}", document_id, source_hash, "logical_row", f"page:{i + 1}:row:{block_no}:{line_no}", "pdf_structure", " ".join(line_words[(block_no, line_no)]), bbox, {"member_ids": members, "page_index": i, "page_rotation": page.rotation, "structure_role": "native_line"}))
        # Reconstruct visual rows independently of PDF text blocks. Columns
        # produced by table generators are often separate PDF lines despite
        # sharing the same baseline. Centre/baseline agreement avoids merging
        # adjacent rows whose glyph rectangles overlap slightly.
        visual_groups: list[dict[str, object]] = []
        native = [o for o in observations if o.channel == "pdf_native" and o.metadata.get("page_index") == i and o.bbox]
        for word in sorted(native, key=lambda o: (((o.bbox[1] + o.bbox[3]) / 2) if o.bbox else 0, o.bbox[0] if o.bbox else 0)):
            assert word.bbox is not None
            selected = None
            for group in reversed(visual_groups):
                if _same_visual_baseline(float(group["center"]), float(group["height"]), word.bbox):
                    selected = group
                    break
            if selected is None:
                selected = {"center": (word.bbox[1] + word.bbox[3]) / 2, "height": word.bbox[3] - word.bbox[1], "words": []}
                visual_groups.append(selected)
            selected["words"].append(word)
            count = len(selected["words"])
            selected["center"] = (float(selected["center"]) * (count - 1) + (word.bbox[1] + word.bbox[3]) / 2) / count
            selected["height"] = min(float(selected["height"]), word.bbox[3] - word.bbox[1])
        for visual_no, group in enumerate(visual_groups):
            row_words = sorted(group["words"], key=lambda o: o.bbox[0] if o.bbox else 0)
            boxes = [o.bbox for o in row_words if o.bbox]
            bbox = (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))
            row_id = f"{document_id}:pdf:p{i + 1}:visual:{visual_no}"
            observations.append(SourceObservation(row_id, document_id, source_hash, "logical_row", f"page:{i + 1}:visual:{visual_no}", "pdf_structure", " ".join(o.text or "" for o in row_words), bbox, {"member_ids": [o.id for o in row_words], "page_index": i, "page_rotation": page.rotation, "structure_role": "visual_row"}))
            observations.extend(_semantic_spans(document_id, source_hash, i, row_id, row_words, "pdf_structure"))
        needs_ocr = not words or bool(images) or bool(drawings and len(words) < 10)
        if needs_ocr:
            if ocr_provider == "azure_document_intelligence" and azure_pages.get(i):
                ocr, issues = azure_pages[i], []
            else:
                provider = OCR_PROVIDERS.get("tesseract" if ocr_provider in {"azure_document_intelligence", "mistral_document_ai"} else ocr_provider) if ocr_enabled else None
                if provider is None:
                    ocr, issues = [], [f"page:{i + 1}:ocr_provider_unavailable:{ocr_provider}" if ocr_enabled else f"page:{i + 1}:ocr_disabled"]
                else:
                    ocr, issues = provider.acquire(page, document_id, source_hash, i, bounds, deadline)
            observations.extend(ocr); diagnostics.extend(issues)
            if ocr_provider == "mistral_document_ai" and i in mistral_pages:
                candidates = _mistral_recognition_candidates(mistral_pages[i], observations, document_id, source_hash, i)
                observations.extend(candidates)
            for word in (item for item in ocr if item.kind == "word"):
                coverage.append(CoverageItem(f"{document_id}:coverage:{word.id}", document_id, "pdf_word", word.locator, CoverageDisposition.UNRESOLVED, reason="awaiting_semantic_disposition"))
            if len(observations) > bounds.max_cells:
                raise ResourceLimitError(f"pdf_observation_limit:{len(observations)}>{bounds.max_cells}")
            if not ocr:
                partial = True
        # Rows of an Azure table belong to that table's own region (one per
        # table and page; tables are not stitched across pages).
        page_rows = [o.id for o in observations if o.metadata.get("page_index") == i and o.kind == "logical_row" and o.metadata.get("structure_role") == "visual_row" and "table_index" not in o.metadata]
        record_blocks = _record_blocks(observations, document_id, source_hash, i, page.rotation)
        observations.extend(record_blocks)
        first_block_top = min((block.bbox[1] for block in record_blocks if block.bbox), default=float("inf"))
        header_rows = [row_id for row_id in page_rows if (row := next((candidate for candidate in observations if candidate.id == row_id), None)) and row.bbox and row.bbox[3] < first_block_top]
        observations.append(SourceObservation(f"{document_id}:pdf:p{i + 1}:region", document_id, source_hash, "table_region", f"page:{i + 1}:region", "pdf_structure", f"Page {i + 1}", metadata={"member_ids": page_rows, "header_row_ids": header_rows, "record_block_ids": [block.id for block in record_blocks], "page_index": i, "page_rotation": page.rotation, "structure_role": "page_region"}))
        ocr_succeeded = any(o.channel == "pdf_ocr" and f"p{i + 1}:" in o.id for o in observations)
        coverage.append(CoverageItem(page_id, document_id, "pdf_page", f"page:{i + 1}", CoverageDisposition.UNRESOLVED if needs_ocr and not ocr_succeeded else CoverageDisposition.OUT_OF_SCOPE, reason="ocr_unresolved" if needs_ocr and not ocr_succeeded else "container_inventory"))
        for j, image in enumerate(images):
            coverage.append(CoverageItem(f"{document_id}:coverage:page:{i + 1}:image:{j}", document_id, "pdf_image", f"page:{i + 1}:image:{j}", CoverageDisposition.OUT_OF_SCOPE if ocr_succeeded or not needs_ocr else CoverageDisposition.UNRESOLVED, reason="ocr_processed_raster_source" if ocr_succeeded else "raster_region_inventory"))
    doc.close()
    return AcquisitionResult(tuple(observations), tuple(coverage), tuple(diagnostics), partial)


def _mistral_recognition_candidates(page: MistralOCRPage, observations: list[SourceObservation], document_id: str, source_hash: str, page_index: int) -> list[SourceObservation]:
    anchors = sorted(
        (
            observation for observation in observations
            if observation.kind == "word"
            and observation.metadata.get("page_index") == page_index
            and observation.channel in {"pdf_native", "pdf_ocr"}
            and observation.bbox
            and observation.text
        ),
        key=lambda observation: ((observation.bbox[1] + observation.bbox[3]) / 2, observation.bbox[0]),
    )
    visual_rows = [
        observation for observation in observations
        if observation.kind == "logical_row"
        and observation.metadata.get("page_index") == page_index
        and observation.metadata.get("structure_role") == "visual_row"
    ]
    rows_by_member: dict[str, list[str]] = {}
    for row in visual_rows:
        for member_id in row.metadata.get("member_ids", ()):
            rows_by_member.setdefault(str(member_id), []).append(row.id)
    provider_sequences = [("page", None, tuple(page.recognition_evidence))]
    provider_sequences.extend(
        ("table", table.id, tuple(table.recognition_evidence))
        for table in getattr(page, "tables", ())
    )
    if not anchors or not any(sequence for _, _, sequence in provider_sequences):
        return []
    anchor_tokens = [norm(observation.text or "") for observation in anchors]
    proposed: list[dict[str, object]] = []
    for provider_scope, provider_table_id, provider in provider_sequences:
        if not provider:
            continue
        provider_tokens = [norm(item.text) for item in provider]
        matches = [block for block in SequenceMatcher(a=anchor_tokens, b=provider_tokens, autojunk=False).get_matching_blocks() if block.size]
        for before, after in zip(matches, matches[1:]):
            anchor_start, anchor_end = before.a + before.size, after.a
            provider_start, provider_end = before.b + before.size, after.b
            gap_length = anchor_end - anchor_start
            if (
                gap_length != provider_end - provider_start
                or gap_length <= 0
                or anchor_start == 0
                or anchor_end >= len(anchors)
                or provider_start == 0
                or provider_end >= len(provider)
            ):
                continue
            left_token = provider_tokens[provider_start - 1]
            right_token = provider_tokens[provider_end]
            matching_contexts = [
                start for start in range(1, len(anchors) - gap_length)
                if anchor_tokens[start - 1] == left_token
                and anchor_tokens[start + gap_length] == right_token
            ]
            # SequenceMatcher can select one of several repeated rows. A
            # recognition candidate is usable only when its two exact anchors
            # identify one physical source occurrence on the page.
            if matching_contexts != [anchor_start]:
                continue
            for anchor_index, provider_index in zip(range(anchor_start, anchor_end), range(provider_start, provider_end)):
                anchor = anchors[anchor_index]; recognition = provider[provider_index]
                if norm(anchor.text or "") == norm(recognition.text):
                    continue
                owner_rows = rows_by_member.get(anchor.id, ())
                if len(owner_rows) != 1:
                    continue
                proposed.append({
                    "anchor": anchor,
                    "recognition": recognition,
                    "provider_scope": provider_scope,
                    "provider_table_id": provider_table_id,
                    "authoritative_row_id": owner_rows[0],
                    "left_anchor_observation_id": anchors[anchor_start - 1].id,
                    "right_anchor_observation_id": anchors[anchor_end].id,
                })

    by_anchor: dict[str, list[dict[str, object]]] = {}
    for item in proposed:
        anchor = item["anchor"]
        assert isinstance(anchor, SourceObservation)
        by_anchor.setdefault(anchor.id, []).append(item)
    candidates: list[SourceObservation] = []
    for anchor_id, items in by_anchor.items():
        texts = {norm(str(item["recognition"].text)) for item in items}
        table_items = [item for item in items if item["provider_scope"] == "table"]
        table_ids = {str(item["provider_table_id"]) for item in table_items if item["provider_table_id"] is not None}
        ownership_closed = len(texts) == 1 and len(table_ids) == 1
        # Collapse duplicate page/table proposals for the same correction, but
        # retain every provider scope and table id as provenance. Conflicting
        # proposals remain separate and unusable so semantics must abstain.
        groups: dict[str, list[dict[str, object]]] = {}
        for item in items:
            groups.setdefault(norm(str(item["recognition"].text)), []).append(item)
        for group in groups.values():
            representative = next((item for item in group if item["provider_scope"] == "table"), group[0])
            anchor = representative["anchor"]
            recognition = representative["recognition"]
            assert isinstance(anchor, SourceObservation)
            provider_table_ids = tuple(sorted({str(item["provider_table_id"]) for item in group if item["provider_table_id"] is not None}))
            candidates.append(SourceObservation(
                f"{document_id}:pdf:p{page_index + 1}:mistral-candidate:{len(candidates)}",
                document_id,
                source_hash,
                "recognition_candidate",
                f"page:{page_index + 1}:mistral-candidate:{len(candidates)}",
                "pdf_recognition_candidate",
                recognition.text,
                anchor.bbox,
                {
                    "page_index": page_index,
                    "anchor_observation_id": anchor_id,
                    "provider": "mistral_document_ai",
                    "provider_confidence": recognition.confidence,
                    "provider_start_index": recognition.start_index,
                    "provider_scopes": tuple(sorted({str(item["provider_scope"]) for item in group})),
                    "provider_table_ids": provider_table_ids,
                    "authoritative_row_id": representative["authoritative_row_id"],
                    "left_anchor_observation_id": representative["left_anchor_observation_id"],
                    "right_anchor_observation_id": representative["right_anchor_observation_id"],
                    "ownership_closed": ownership_closed,
                    "alignment": "two_sided_exact_sequence",
                    "coordinate_authority": "native_or_tesseract_anchor",
                    "candidate_only": True,
                    "cell_structure_authoritative": False,
                },
            ))
    return candidates


def _record_blocks(observations: list[SourceObservation], document_id: str, source_hash: str, page_index: int, rotation: int, *, table_index: int | None = None) -> list[SourceObservation]:
    """Group measure rows into record blocks.

    With ``table_index`` None this covers the page's word rows; otherwise only
    the rows of that Azure table, whose members are ``cell`` observations.
    """
    by_id = {observation.id: observation for observation in observations}
    rows = sorted(
        (
            observation for observation in observations
            if observation.metadata.get("page_index") == page_index and observation.kind == "logical_row" and observation.metadata.get("structure_role") == "visual_row" and observation.bbox
            and observation.metadata.get("table_index") == table_index
        ),
        key=lambda observation: observation.bbox[1],
    )
    scope_members = {str(member) for row in rows for member in row.metadata.get("member_ids", ())}
    group_spans = [
        observation for observation in observations
        if observation.metadata.get("page_index") == page_index and observation.metadata.get("semantic_group") and observation.bbox
        and scope_members.intersection(str(member) for member in observation.metadata.get("member_ids", ()))
    ]
    if not group_spans:
        return []
    value_boundary = min(span.bbox[0] for span in group_spans)
    candidates: list[tuple[SourceObservation, SourceObservation | None, bool]] = []
    for row in rows:
        row_members = {str(value) for value in row.metadata.get("member_ids", ())}
        spans = [span for span in group_spans if row_members.intersection(str(value) for value in span.metadata.get("member_ids", ()))]
        span = spans[0] if len(spans) == 1 else None
        words = [by_id[value] for value in row_members if value in by_id and by_id[value].kind in {"word", "cell"} and by_id[value].bbox]
        has_financial_value = False
        for word in words:
            if (word.bbox[0] + word.bbox[2]) / 2 <= value_boundary or not word.text:
                continue
            try:
                parse_amount(word.text); has_financial_value = True; break
            except NumericParseError:
                continue
        left_words = sorted((word for word in words if word.bbox[2] < value_boundary), key=lambda word: word.bbox[0])
        leading_identifier = bool(left_words and re.fullmatch(r"(?=.*\d)[A-Za-z0-9._/-]+", left_words[0].text or ""))
        if has_financial_value:
            candidates.append((row, span, leading_identifier))
    blocks: list[SourceObservation] = []
    current: list[tuple[SourceObservation, SourceObservation | None]] = []
    parent_row_id: str | None = None
    candidate_rows = {row.id for row, _, _ in candidates}
    preceding_nonmeasure = [row for row in rows if row.id not in candidate_rows]

    def emit() -> None:
        nonlocal current
        if not current:
            return
        block_rows = [item[0] for item in current]; spans = [item[1] for item in current if item[1] is not None]
        boxes = [row.bbox for row in block_rows if row.bbox]
        bbox = (min(box[0] for box in boxes), min(box[1] for box in boxes), max(box[2] for box in boxes), max(box[3] for box in boxes))
        scope = "record" if table_index is None else f"aztable:{table_index}:record"
        block_id = f"{document_id}:pdf:p{page_index + 1}:{scope}:{len(blocks)}"
        blocks.append(SourceObservation(block_id, document_id, source_hash, "record_block", f"page:{page_index + 1}:{scope}:{len(blocks)}", "pdf_structure", " | ".join(row.text or "" for row in block_rows), bbox, {
            "member_ids": [row.id for row in block_rows],
            "identity_row_id": block_rows[0].id,
            "measure_row_ids": [row.id for row in block_rows],
            "measure_span_ids": [span.id for span in spans],
            "semantic_measures": [span.metadata.get("semantic_group") if span else None for _, span in current],
            "parent_row_id": parent_row_id,
            "page_index": page_index,
            "page_rotation": rotation,
            "structure_role": "record_block",
            **({} if table_index is None else {"table_index": table_index}),
        }))
        current = []

    seen_measures: set[str] = set()
    previous_candidate_index: int | None = None
    row_indexes = {row.id: index for index, row in enumerate(rows)}
    for row, span, starts_record in candidates:
        measure = str(span.metadata.get("semantic_group")) if span is not None else None
        separated = previous_candidate_index is not None and row_indexes[row.id] != previous_candidate_index + 1
        repeated_measure = measure is not None and measure in seen_measures
        if separated or repeated_measure:
            emit(); seen_measures = set()
        if starts_record:
            emit()
            parent = [candidate for candidate in preceding_nonmeasure if candidate.bbox and candidate.bbox[3] < row.bbox[1]]
            parent_row_id = parent[-1].id if parent else None
            current = [(row, span)]
        elif current:
            current.append((row, span))
        if current and measure is not None:
            seen_measures.add(measure)
        previous_candidate_index = row_indexes[row.id]
    emit()
    return blocks


def _azure_page_observations(provider_page: OCRPage, source_page: fitz.Page, document_id: str, source_hash: str, page_index: int, tables: tuple[OCRTable, ...] = ()) -> tuple[list[SourceObservation], list[str]]:
    if provider_page.unit not in {"inch", "pixel"} or provider_page.width <= 0 or provider_page.height <= 0:
        raise ValueError(f"unsupported_azure_coordinate_unit:{provider_page.unit}")
    canonical = source_page.cropbox
    rotation = source_page.rotation % 360
    rendered_width, rendered_height = ((canonical.height, canonical.width) if rotation in {90, 270} else (canonical.width, canonical.height))
    transform = PageTransform(rotation=rotation, page_width=canonical.width, page_height=canonical.height)

    # Azure's page angle is the clockwise orientation of the content. Whole
    # quarter turns are the page /Rotate (undone by PageTransform); the
    # remainder is scan skew, undone here about the page centre so pointers
    # land in the upright page's coordinates.
    deskew_degrees = 0.0 if provider_page.angle is None else ((provider_page.angle + 45.0) % 90.0) - 45.0
    cos_t, sin_t = math.cos(math.radians(deskew_degrees)), math.sin(math.radians(deskew_degrees))
    centre_x, centre_y = provider_page.width / 2, provider_page.height / 2

    def to_source(polygon: tuple[float, ...], *, deskew: bool = True) -> tuple[tuple[float, float, float, float], tuple[float, float, float, float]]:
        # Azure polygons are in the displayed (rotated, cropped) page in
        # provider units; deskew, scale to points, then undo the page rotation.
        # deskew=False gives the raw printed position on the scan.
        upright = []
        for offset in range(0, len(polygon), 2):
            dx, dy = polygon[offset] - centre_x, polygon[offset + 1] - centre_y
            upright.append((centre_x + dx * cos_t + dy * sin_t, centre_y - dx * sin_t + dy * cos_t) if deskew else (polygon[offset], polygon[offset + 1]))
        points = tuple((x / provider_page.width * rendered_width, y / provider_page.height * rendered_height) for x, y in upright)
        rendered = (min(x for x, _ in points), min(y for _, y in points), max(x for x, _ in points), max(y for _, y in points))
        return rendered, transform.to_source(rendered)

    observations: list[SourceObservation] = []
    for index, word in enumerate(provider_page.words):
        rendered, source_box = to_source(word.polygon)
        oid = f"{document_id}:pdf:p{page_index + 1}:azure:{index}"
        observations.append(SourceObservation(oid, document_id, source_hash, "word", f"page:{page_index + 1}:azure:{index}", "pdf_ocr", word.text, source_box, {"confidence": word.confidence, "provider": "azure_document_intelligence", "provider_unit": provider_page.unit, "provider_page_width": provider_page.width, "provider_page_height": provider_page.height, "render_box": rendered, "deskew_degrees": deskew_degrees, "pointer_convention": POINTER_CONVENTION, "raw_bbox": to_source(word.polygon, deskew=False)[1], "roundtrip_ok": transform.roundtrip_ok(source_box), "page_index": page_index, "page_rotation": rotation, "page_width": canonical.width, "page_height": canonical.height}))
    groups: list[list[SourceObservation]] = []
    for word in sorted(observations, key=lambda item: (item.bbox[1], item.bbox[0]) if item.bbox else (0, 0)):
        if not word.bbox:
            continue
        group = next((candidate for candidate in reversed(groups) if candidate[0].bbox and _same_visual_baseline((candidate[0].bbox[1] + candidate[0].bbox[3]) / 2, candidate[0].bbox[3] - candidate[0].bbox[1], word.bbox)), None)
        if group is None:
            group = []; groups.append(group)
        group.append(word)
    for row_no, group in enumerate(groups):
        group.sort(key=lambda item: item.bbox[0] if item.bbox else 0); boxes = [item.bbox for item in group if item.bbox]
        bbox = (min(box[0] for box in boxes), min(box[1] for box in boxes), max(box[2] for box in boxes), max(box[3] for box in boxes))
        row_id = f"{document_id}:pdf:p{page_index + 1}:azurerow:{row_no}"
        observations.append(SourceObservation(row_id, document_id, source_hash, "logical_row", f"page:{page_index + 1}:azurerow:{row_no}", "pdf_ocr_structure", " ".join(item.text or "" for item in group), bbox, {"member_ids": [item.id for item in group], "page_index": page_index, "page_rotation": rotation, "structure_role": "visual_row", "provider": "azure_document_intelligence"}))
        # Narrowest span per (role, value), as for native rows. Emitting every
        # matching span keyed only by its start produced duplicate ids, which
        # the ledger rejects (immutable record collision).
        observations.extend(_semantic_spans(document_id, source_hash, page_index, row_id, group, "pdf_ocr_structure"))
    page_words = [(f"{document_id}:pdf:p{page_index + 1}:azure:{index}", word) for index, word in enumerate(provider_page.words)]
    cell_observations, issues = _azure_table_observations(provider_page, page_words, tables, document_id, source_hash, page_index, transform, canonical, to_source, deskew_degrees)
    observations.extend(cell_observations)
    return observations, issues


def _union(boxes: list[tuple[float, float, float, float]]) -> tuple[float, float, float, float] | None:
    if not boxes:
        return None
    return (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))


def _normalise_space(text: str) -> str:
    return " ".join(text.split())


_NUMERIC_LOOKING = re.compile(r"[\d\s.,'+\-\u2212()]*\d[\d\s.,'+\-\u2212()]*")


def _point_in_polygon(x: float, y: float, polygon: tuple[float, ...]) -> bool:
    inside = False
    points = [(polygon[offset], polygon[offset + 1]) for offset in range(0, len(polygon), 2)]
    for (x0, y0), (x1, y1) in zip(points, points[1:] + points[:1]):
        if (y0 > y) != (y1 > y) and x < x0 + (y - y0) * (x1 - x0) / (y1 - y0):
            inside = not inside
    return inside


# Azure builds a cell's text from the same words the text check compares, so a
# misread digit passes that check; word confidence is the only signal left. The
# clean recorded pages (synthetic scans, f01) have every word at >= 0.988; the
# real f05 SAP page has 95 of 221 numeric words below 0.9, its misreads among
# them ("1.456,941,12-" at 0.601). Set by the code review, not yet calibrated
# on the corpus.
MIN_NUMERIC_WORD_CONFIDENCE = 0.9

# Pointers of Azure observations are in the upright (deskewed) page; the raw
# printed position is kept in metadata (raw_bbox, raw_cell_polygon, raw_lexeme_bbox).
POINTER_CONVENTION = "deskewed_upright"


def _cross_check(content: str, words: list[OCRWord], regions_on_page: int, shared: bool) -> str:
    """Return "verified" or why the cell text is not corroborated by its own words."""
    if regions_on_page != 1:
        return "multiple_regions"
    if shared:
        return "word_in_multiple_cells"
    joined = _normalise_space(" ".join(word.text for word in words))
    text = _normalise_space(content)
    if text != joined:
        return "text_mismatch"
    if text and _NUMERIC_LOOKING.fullmatch(text):
        if any(word.confidence < MIN_NUMERIC_WORD_CONFIDENCE for word in words):
            return "low_confidence"
        try:
            if parse_amount(text).source_signed_value != parse_amount(joined).source_signed_value:
                return "numeric_mismatch"
        except NumericParseError:
            return "numeric_unparseable"
    return "verified"


def _azure_table_observations(provider_page: OCRPage, page_words: list[tuple[str, OCRWord]], tables: tuple[OCRTable, ...], document_id: str, source_hash: str, page_index: int, transform: PageTransform, canonical: fitz.Rect, to_source: Callable[..., tuple[tuple[float, float, float, float], tuple[float, float, float, float]]], deskew_degrees: float = 0.0) -> tuple[list[SourceObservation], list[str]]:
    """Azure table cells as `cell` observations, with rows and one region per table and page.

    Cell text is kept as Azure returned it. A cell is `verified` only when the
    Azure words whose centre lies inside its polygon spell the same text
    (whitespace-normalised) and, for numeric text, parse to the same amount.
    Tables are not stitched across pages.
    """
    page_number = provider_page.page_number
    rotation = transform.rotation
    observations: list[SourceObservation] = []
    verified_count = unverified_count = 0
    word_centres = []
    for word_id, word in page_words:
        xs, ys = word.polygon[0::2], word.polygon[1::2]
        word_centres.append((word_id, word, (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2))
    placed_by_table: list[list[tuple[OCRTableCell, tuple[float, ...], int, list[tuple[str, OCRWord]]]]] = []
    for table in tables:
        placed: list[tuple[OCRTableCell, tuple[float, ...], int, list[tuple[str, OCRWord]]]] = []
        for cell in table.cells:
            regions = [region for region in cell.bounding_regions if region.page_number == page_number]
            if not regions:
                continue
            polygon = regions[0].polygon
            members = sorted(((word_id, word) for word_id, word, x, y in word_centres if _point_in_polygon(x, y, polygon)), key=lambda item: item[1].span.offset)
            placed.append((cell, polygon, len(regions), members))
        placed_by_table.append(placed)
    # Claims are counted across all tables on the page: a word inside cells of
    # two overlapping tables must not verify in both.
    claims: dict[str, int] = {}
    for placed in placed_by_table:
        for _, _, _, members in placed:
            for word_id, _ in members:
                claims[word_id] = claims.get(word_id, 0) + 1
    for table_index, (table, placed) in enumerate(zip(tables, placed_by_table)):
        if not placed:
            continue
        cells_by_row: dict[int, list[SourceObservation]] = {}
        single_row_boxes: dict[int, list[tuple[float, float, float, float]]] = {}
        for cell, polygon, region_count, members in placed:
            rendered, source_box = to_source(polygon)
            status = _cross_check(cell.content, [word for _, word in members], region_count, any(claims[word_id] > 1 for word_id, _ in members))
            verified = status == "verified"
            lexeme_bbox = _union([to_source(word.polygon)[1] for _, word in members]) if verified else None
            raw_lexeme_bbox = _union([to_source(word.polygon, deskew=False)[1] for _, word in members]) if verified else None
            verified_count += verified
            unverified_count += not verified
            row_span, column_span = cell.row_span or 1, cell.column_span or 1
            locator = f"page:{page_index + 1}:azcell:{table_index}:{cell.row_index}:{cell.column_index}"
            observation = SourceObservation(
                f"{document_id}:pdf:p{page_index + 1}:azcell:{table_index}:{cell.row_index}:{cell.column_index}", document_id, source_hash, "cell",
                locator, "pdf_ocr", cell.content, source_box,
                {
                    "row_index": cell.row_index, "column_index": cell.column_index, "row_span": row_span, "column_span": column_span,
                    "cell_kind": cell.kind or "content", "table_index": table_index, "page_index": page_index, "page_rotation": rotation,
                    "page_width": canonical.width, "page_height": canonical.height,
                    "confidence": min((word.confidence for _, word in members), default=None),
                    "min_word_confidence": min((word.confidence for _, word in members), default=None),
                    "azure_cell_id": f"t{table_index}r{cell.row_index}c{cell.column_index}", "verified": verified, "cross_check": status,
                    # The printed lexeme inside the cell (union of its verified
                    # words); the cell polygon itself is usually wider.
                    "lexeme_bbox": lexeme_bbox,
                    # Kept beside the lexeme box so the export policy can switch
                    # the pointer convention once the official scorer is known.
                    "cell_polygon": source_box,
                    "pointer_convention": POINTER_CONVENTION,
                    "raw_cell_polygon": to_source(polygon, deskew=False)[1],
                    "raw_lexeme_bbox": raw_lexeme_bbox,
                    "member_word_ids": [word_id for word_id, _ in members],
                    "provider": "azure_document_intelligence", "render_box": rendered, "deskew_degrees": deskew_degrees, "roundtrip_ok": transform.roundtrip_ok(source_box),
                },
            )
            observations.append(observation)
            cells_by_row.setdefault(cell.row_index, []).append(observation)
            if row_span == 1:
                single_row_boxes.setdefault(cell.row_index, []).append(source_box)
        row_ids: list[str] = []
        header_row_ids: list[str] = []
        table_observations: list[SourceObservation] = []
        for row_index in sorted(cells_by_row):
            # A spanning cell belongs only to the row it starts in, so every
            # cell has exactly one owning visual row.
            members = sorted(cells_by_row[row_index], key=lambda item: int(item.metadata["column_index"]))
            # A cell spanning several rows would stretch the row box over its
            # neighbours; prefer the boxes of the row's single-row cells.
            boxes = single_row_boxes.get(row_index) or [item.bbox for item in members if item.bbox]
            bbox = (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))
            row_id = f"{document_id}:pdf:p{page_index + 1}:azrow:{table_index}:{row_index}"
            table_observations.append(SourceObservation(row_id, document_id, source_hash, "logical_row", f"page:{page_index + 1}:azrow:{table_index}:{row_index}", "pdf_ocr_structure", " ".join(item.text or "" for item in members), bbox, {"member_ids": [item.id for item in members], "page_index": page_index, "page_rotation": rotation, "structure_role": "visual_row", "provider": "azure_document_intelligence", "table_index": table_index, "row_index": row_index}))
            row_ids.append(row_id)
            if any(item.metadata["cell_kind"] == "columnHeader" for item in members):
                header_row_ids.append(row_id)
            table_observations.extend(_semantic_spans(document_id, source_hash, page_index, row_id, members, "pdf_ocr_structure"))
        blocks = _record_blocks([*observations, *table_observations], document_id, source_hash, page_index, rotation, table_index=table_index)
        table_observations.extend(blocks)
        regions = [region.polygon for region in table.bounding_regions if region.page_number == page_number]
        table_observations.append(SourceObservation(
            f"{document_id}:pdf:p{page_index + 1}:aztable:{table_index}", document_id, source_hash, "table_region", f"page:{page_index + 1}:aztable:{table_index}", "pdf_ocr_structure", f"Page {page_index + 1} table {table_index}", to_source(regions[0])[1] if len(regions) == 1 else None,
            {"member_ids": row_ids, "header_row_ids": header_row_ids, "record_block_ids": [block.id for block in blocks], "page_index": page_index, "page_rotation": rotation, "structure_role": "azure_table", "provider": "azure_document_intelligence", "table_index": table_index, "row_count": table.row_count, "column_count": table.column_count},
        ))
        observations.extend(table_observations)
    if not verified_count and not unverified_count:
        return observations, []
    return observations, [f"page:{page_index + 1}:azure_cells_verified:{verified_count}", f"page:{page_index + 1}:azure_cells_unverified:{unverified_count}"]

