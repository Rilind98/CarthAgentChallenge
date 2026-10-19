from __future__ import annotations

from dataclasses import asdict, dataclass, fields, replace
from hashlib import sha256
import json
import multiprocessing
import fcntl
from pathlib import Path
import resource
import sys
import time
import tomllib
from typing import Any

from .acquire_pdf import acquire_pdf
from .acquire_xlsx import AcquisitionResult, ResourceLimitError, acquire_xlsx
from .exporter import SCHEMA_VERSION, Snapshot, TemplateExportAdapter, publish
from .mapper import MAPPER_PROMPT_VERSION, MAPPER_SCHEMA_VERSION, MapperConfig, MapperUnavailableError, MapperValidationError, OpenAICompatibleMapper
from .ocr_azure import AzureDocumentIntelligenceOCR, AzureOCRConfig
from .ocr_mistral import MistralDocumentOCR, MistralOCRConfig
from .models import AssociationClaim, AssociationKind, CoverageDisposition, CoverageItem, HypothesisSet, OutputBinding, OutputIdentity, RunStatus, SourceObservation, jsonable
from .policies import DEFAULT_BOUNDS, DEFAULT_POLICIES, PolicyBundle, ResourceBounds
from .semantics import SEMANTICS_VERSION, interpret
from . import semantics as semantics_module
from .store import Ledger, content_hash, preserve_source


PIPELINE_VERSION = "0.13.1"
ACQUISITION_VERSION = "15"
STRUCTURE_VERSION = "9"


def _implementation_fingerprint() -> dict[str, str]:
    return {
        "pipeline": PIPELINE_VERSION,
        "acquisition": ACQUISITION_VERSION,
        "structure": STRUCTURE_VERSION,
        "semantics": SEMANTICS_VERSION,
        "mapper_prompt": MAPPER_PROMPT_VERSION,
        "mapper_schema": MAPPER_SCHEMA_VERSION,
        "evidence_schema": SCHEMA_VERSION,
        "template_adapter": TemplateExportAdapter.CONTRACT_VERSION,
        "template_adapter_implementation": TemplateExportAdapter.IMPLEMENTATION_VERSION,
    }


@dataclass(frozen=True, slots=True)
class OCRConfig:
    enabled: bool = True
    provider: str = "tesseract"
    azure: AzureOCRConfig | None = None
    mistral: MistralOCRConfig | None = None


@dataclass(frozen=True, slots=True)
class RunConfig:
    bounds: ResourceBounds = DEFAULT_BOUNDS
    policies: PolicyBundle = DEFAULT_POLICIES
    mapper: MapperConfig = MapperConfig()
    template: Path | None = None
    ocr: OCRConfig = OCRConfig()


@dataclass(frozen=True, slots=True)
class RunResult:
    status: RunStatus
    run_id: str
    document_id: str
    source_hash: str | None
    output_dir: Path | None
    evidence_path: Path | None
    review_path: Path | None
    manifest_path: Path | None
    diagnostics: tuple[str, ...]
    duration_seconds: float


def load_config(path: Path | None, template: Path | None = None) -> RunConfig:
    if path is None:
        return RunConfig(template=template)
    raw = path.read_bytes()
    data = json.loads(raw) if path.suffix.lower() == ".json" else tomllib.loads(raw.decode())
    allowed_bounds = {f.name for f in fields(ResourceBounds)}
    bounds_data = data.get("bounds", {})
    unknown = set(bounds_data) - allowed_bounds
    if unknown:
        raise ValueError(f"unknown bound(s): {sorted(unknown)}")
    mapper_data = data.get("mapper", {})
    ocr_data = dict(data.get("ocr", {}))
    azure_data = ocr_data.pop("azure", None)
    azure = AzureOCRConfig(**azure_data) if azure_data is not None else None
    mistral_data = ocr_data.pop("mistral", None)
    mistral = MistralOCRConfig(**mistral_data) if mistral_data is not None else None
    return RunConfig(ResourceBounds(**{**asdict(DEFAULT_BOUNDS), **bounds_data}), DEFAULT_POLICIES, MapperConfig(**mapper_data), template or (Path(data["template"]) if data.get("template") else None), OCRConfig(**ocr_data, azure=azure, mistral=mistral))


def _document_id(path: Path, source_hash: str | None = None) -> str:
    clean = "".join(c if c.isalnum() or c in "-_" else "-" for c in path.stem).strip("-") or "document"
    clean = clean[:67]
    return f"{clean}-{source_hash[:12]}" if source_hash else clean


def run_input(input_path: Path, output: Path, config: RunConfig = RunConfig()) -> RunResult:
    started = time.monotonic(); document_id = _document_id(input_path)
    if input_path.suffix.lower() in {".doc", ".docx"}:
        return RunResult(RunStatus.UNSUPPORTED, f"unsupported:{document_id}", document_id, None, None, None, None, None, ("unsupported_input:doc_docx",), time.monotonic() - started)
    if input_path.suffix.lower() not in {".pdf", ".xlsx"}:
        return RunResult(RunStatus.FAILED, f"failed:{document_id}", document_id, None, None, None, None, None, ("unsupported_extension",), time.monotonic() - started)
    try:
        size = input_path.stat().st_size
        if size > config.bounds.max_input_bytes:
            raise ResourceLimitError(f"input_size_limit:{size}>{config.bounds.max_input_bytes}")
        source_hash = content_hash(input_path)
        document_id = _document_id(input_path, source_hash)
        lock_dir = output / ".carthagent" / "locks"; lock_dir.mkdir(parents=True, exist_ok=True)
        document_lock = (lock_dir / f"{document_id}.lock").open("a+b")
        fcntl.flock(document_lock.fileno(), fcntl.LOCK_EX)
        mapper = OpenAICompatibleMapper(config.mapper)
        mapper_fingerprint = {**mapper.fingerprint(), "chunking": {"pdf_pages": config.bounds.pdf_chunk_pages, "pdf_overlap": config.bounds.page_overlap, "logical_rows": config.bounds.row_chunk_size, "row_overlap": config.bounds.row_overlap}}
        template_hash = content_hash(config.template) if config.template else None
        implementation = _implementation_fingerprint()
        run_id = sha256(f"{json.dumps(implementation,sort_keys=True)}:{source_hash}:{config.policies.digest()}:{json.dumps(asdict(config.bounds),sort_keys=True)}:{json.dumps(mapper_fingerprint,sort_keys=True)}:{json.dumps(asdict(config.ocr),sort_keys=True)}:{template_hash}".encode()).hexdigest()[:24]
        ledger = Ledger(output / ".carthagent" / "ledger.sqlite3", busy_timeout_seconds=max(0.0, config.bounds.document_timeout_seconds - (time.monotonic() - started)))
        source_copy = preserve_source(input_path, output / ".carthagent" / "sources", source_hash)
        ledger.append(f"source:{source_hash}", document_id, "source_bytes:2", {"source_hash": source_hash, "size": size})
        final = output / document_id
        prior = ledger.get_checkpoint(document_id, "publication_verified")
        if prior and prior.get("export_record_id") and not ledger.is_invalidated(prior["export_record_id"]) and final.exists() and (final / ".complete").exists():
            try:
                manifest = json.loads((final / "manifest.v1.json").read_text())
                if manifest["run_id"] == run_id and manifest["source_hash"] == source_hash and _verified_artifacts(final, manifest, prior.get("manifest_hash")):
                    ledger.close()
                    return RunResult(RunStatus(manifest["status"]), run_id, document_id, source_hash, final, final / "evidence.v1.json", final / "review.v1.xlsx", final / "manifest.v1.json", ("resumed_verified_publication",), time.monotonic() - started)
            except (OSError, KeyError, ValueError, json.JSONDecodeError):
                pass
        ledger.checkpoint(document_id, "registered", {"run_id": run_id, "source_hash": source_hash})
        deadline = started + config.bounds.document_timeout_seconds
        cache_key = Ledger.cache_key(source_hash, f"whole_document_acquisition:{document_id}", {**asdict(config.bounds), "ocr": asdict(config.ocr)}, {"acquisition": ACQUISITION_VERSION, "structure": STRUCTURE_VERSION})
        cached = ledger.cache_get(cache_key)
        if cached:
            acquired = AcquisitionResult(tuple(SourceObservation(o["id"], o["document_id"], o["source_hash"], o["kind"], o["locator"], o["channel"], o.get("text"), tuple(o["bbox"]) if o.get("bbox") else None, o.get("metadata", {})) for o in cached["observations"]), tuple(CoverageItem(c["id"], c["document_id"], c["kind"], c["locator"], CoverageDisposition(c["disposition"]), c.get("linked_record_id"), c.get("reason")) for c in cached["coverage"]), tuple(cached["diagnostics"]), bool(cached["partial"]))
        else:
            azure_provider = AzureDocumentIntelligenceOCR(config.ocr.azure) if config.ocr.provider == "azure_document_intelligence" and config.ocr.azure is not None else None
            mistral_provider = MistralDocumentOCR(config.ocr.mistral) if config.ocr.provider == "mistral_document_ai" and config.ocr.mistral is not None else None
            acquired = acquire_pdf(source_copy, document_id, source_hash, config.bounds, ocr_provider=config.ocr.provider, ocr_enabled=config.ocr.enabled, deadline=deadline, azure_provider=azure_provider, mistral_provider=mistral_provider) if input_path.suffix.lower() == ".pdf" else acquire_xlsx(source_copy, document_id, source_hash, config.bounds, deadline=deadline)
            ledger.cache_put(cache_key, {"observations": [jsonable(o) for o in acquired.observations], "coverage": [jsonable(c) for c in acquired.coverage], "diagnostics": list(acquired.diagnostics), "partial": acquired.partial})
        if time.monotonic() - started > config.bounds.document_timeout_seconds:
            raise ResourceLimitError(f"document_timeout_after_acquisition:{config.bounds.document_timeout_seconds}s")
        _check_worker_memory(config.bounds)
        ledger.checkpoint(document_id, "acquired", {"observation_count": len(acquired.observations), "coverage_count": len(acquired.coverage)})
        mapper_diagnostics: tuple[str, ...] = ()
        mapped_hypotheses: tuple[HypothesisSet, ...] = ()
        if config.mapper.enabled:
            mapped_chunks: list[HypothesisSet] = []
            diagnostic_items: list[str] = []
            residual_ids = {candidate for abstention in interpret(acquired.observations, acquired.coverage, config.policies, deadline=deadline)[2] for candidate in abstention.candidate_observation_ids}
            mapper_phase_start = time.monotonic()
            mapper_deadline = min(
                mapper_phase_start + config.bounds.mapper_phase_timeout_seconds,
                deadline - config.bounds.post_mapper_reserve_seconds,
            )
            http_attempts = 0
            stop_mapper = mapper_deadline <= mapper_phase_start
            if stop_mapper:
                diagnostic_items.append("mapper_budget_exhausted:post_mapper_reserve")
            for chunk in _observation_chunks(acquired.observations, config.bounds):
                if stop_mapper:
                    break
                for request_chunk in _mapper_bounded_chunks(mapper, document_id, chunk, config.bounds.row_overlap):
                    request_chunk = _regions_needing_mapper(request_chunk, residual_ids)
                    if not request_chunk:
                        continue
                    request_fingerprint = mapper.request_fingerprint(document_id, request_chunk)
                    chunk_cache_key = Ledger.cache_key(
                        source_hash,
                        f"semantic_mapper_chunk:{document_id}:{request_fingerprint}",
                        mapper_fingerprint,
                        {"semantics": SEMANTICS_VERSION},
                    )
                    cached_chunk = ledger.cache_get(chunk_cache_key)
                    if cached_chunk and cached_chunk.get("request_fingerprint") == request_fingerprint:
                        mapped_chunks.extend(_cached_hypotheses(cached_chunk.get("hypotheses", [])))
                        continue
                    if time.monotonic() >= mapper_deadline or http_attempts >= config.bounds.max_mapper_requests:
                        diagnostic_items.append(f"mapper_budget_exhausted:attempts={http_attempts}")
                        stop_mapper = True
                        break
                    remaining_attempts = config.bounds.max_mapper_requests - http_attempts
                    retries = min(config.bounds.transient_retries, remaining_attempts - 1)
                    try:
                        completed = mapper.map(
                            document_id,
                            request_chunk,
                            transient_retries=retries,
                            deadline=mapper_deadline,
                            max_input_bytes=config.bounds.max_input_bytes,
                            max_response_bytes=config.bounds.max_mapper_response_bytes,
                        )
                        http_attempts += mapper.last_http_attempts
                        mapped_chunks.extend(completed)
                        ledger.cache_put(chunk_cache_key, {
                            "request_fingerprint": request_fingerprint,
                            "audit": mapper.last_audit,
                            "hypotheses": [jsonable(hypothesis) for hypothesis in completed],
                        })
                    except (MapperUnavailableError, MapperValidationError) as exc:
                        http_attempts += mapper.last_http_attempts
                        diagnostic_items.append(f"mapper_unavailable:{type(exc).__name__}:{exc}")
                        stop_mapper = True
                        break
            mapped_hypotheses = _deduplicate_mapped(tuple(mapped_chunks))
            mapper_diagnostics = tuple(dict.fromkeys(diagnostic_items))
        hypotheses, facts, abstentions, coverage = interpret(acquired.observations, acquired.coverage, config.policies, mapped_hypotheses, deadline=deadline)
        if time.monotonic() - started > config.bounds.document_timeout_seconds:
            raise ResourceLimitError(f"document_timeout_after_semantics:{config.bounds.document_timeout_seconds}s")
        _check_worker_memory(config.bounds)
        bindings = tuple(OutputBinding(f"{document_id}:binding:{i}", "pending", fact.id, OutputIdentity(config.policies.output_contract.version), fact.source_pointer, "contract_pending") for i, fact in enumerate(facts))
        observation_payload = tuple(jsonable(o) for o in acquired.observations)
        planned = TemplateExportAdapter().plan_bindings("pending", facts) if config.template else bindings
        semantic_snapshot = Snapshot("pending", run_id, document_id, source_hash, config.policies.digest(), facts, abstentions, coverage, observation_payload, planned)
        snapshot_id = semantic_snapshot.digest[:24]
        bindings = tuple(_finalize_binding_id(b, snapshot_id, document_id) for b in planned)
        output_diagnostics = tuple(f"output_coverage:{binding.status}:{binding.record_id}" for binding in bindings if config.template and binding.status != "bound")
        context_diagnostics = tuple(f"semantic_context_missing:period:{fact.id}" for fact in facts if fact.business.period is None)
        status = RunStatus.PARTIAL if acquired.partial or mapper_diagnostics or abstentions or output_diagnostics or context_diagnostics or any(c.disposition.value == "unresolved_unsupported" for c in coverage) else RunStatus.COMPLETE
        snapshot = Snapshot(snapshot_id, run_id, document_id, source_hash, config.policies.digest(), facts, abstentions, coverage, observation_payload, bindings)
        with ledger.transaction():
            for obs in acquired.observations:
                ledger.append(obs.id, document_id, "observation:9", jsonable(obs))
            persisted_hypotheses = {hyp.id: hyp for hyp in (*mapped_hypotheses, *hypotheses)}
            for hyp in persisted_hypotheses.values():
                ledger.append(hyp.id, document_id, "hypothesis", jsonable(hyp))
            for fact in facts:
                ledger.append(fact.id, document_id, "fact", jsonable(fact))
                for dependency in fact.dependency_ids:
                    ledger.depend(fact.id, dependency, "semantic")
            for abst in abstentions:
                ledger.append(abst.id, document_id, "abstention", jsonable(abst))
            for item in coverage:
                ledger.append(item.id, document_id, f"coverage:{PIPELINE_VERSION}:{semantics_module.SEMANTICS_VERSION}", jsonable(item))  # R7-3: a semantics change is a new kind
            ledger.append(snapshot.id, document_id, "snapshot", {"digest": snapshot.digest, "fact_ids": [f.id for f in facts], "abstention_ids": [a.id for a in abstentions], "coverage_ids": [c.id for c in coverage]})
            for fact in facts:
                ledger.depend(snapshot.id, fact.id, "semantic_snapshot")
            for binding in bindings:
                ledger.append(binding.id, document_id, "output_binding", jsonable(binding))  # ids are content-addressed (_finalize_binding_id)
                ledger.depend(binding.id, binding.record_id, "output_projection")
                ledger.depend(binding.id, snapshot.id, "snapshot_projection")
        diagnostics = (*acquired.diagnostics, *mapper_diagnostics, *output_diagnostics, *context_diagnostics)
        final, evidence, review, manifest = publish(snapshot, output, status.value, diagnostics, template=config.template)
        _check_worker_memory(config.bounds)
        manifest_hash = content_hash(manifest)
        export_record_id = f"{document_id}:export:{sha256((snapshot.id + snapshot.digest + manifest_hash).encode()).hexdigest()[:20]}"
        with ledger.transaction():
            ledger.append(export_record_id, document_id, "export", {"snapshot_id": snapshot.id, "snapshot_digest": snapshot.digest, "manifest_hash": manifest_hash, "output_dir": str(final)})
            for binding in bindings:
                ledger.depend(export_record_id, binding.id, "publication")
            ledger.event(export_record_id, "revalidated", {"snapshot_digest": snapshot.digest})
            ledger.checkpoint(document_id, "publication_verified", {"snapshot_id": snapshot.id, "export_record_id": export_record_id, "manifest_hash": manifest_hash, "output_dir": str(final)})
        ledger.close()
        return RunResult(status, run_id, document_id, source_hash, final, evidence, review, manifest, diagnostics, time.monotonic() - started)
    except Exception as exc:
        return RunResult(RunStatus.FAILED, f"failed:{document_id}", document_id, locals().get("source_hash"), None, None, None, None, (f"{type(exc).__name__}:{exc}",), time.monotonic() - started)


def run_many(input_path: Path, output: Path, config: RunConfig) -> tuple[RunResult, ...]:
    if input_path.is_file():
        files = [input_path]
    else:
        output_resolved = output.resolve()
        files = sorted(p for p in input_path.rglob("*") if p.is_file() and p.suffix.lower() in {".pdf", ".xlsx", ".doc", ".docx"} and output_resolved not in p.resolve().parents and not any(part.startswith(("_", ".")) for part in p.relative_to(input_path).parts[:-1]))
    results: list[RunResult | None] = [None] * len(files)
    context = multiprocessing.get_context("spawn")
    for batch_start in range(0, len(files), config.bounds.worker_processes):
        active = []
        for index in range(batch_start, min(batch_start + config.bounds.worker_processes, len(files))):
            receive, send = context.Pipe(duplex=False)
            process = context.Process(target=_bounded_run_worker, args=(send, files[index], output, config), daemon=True)
            process.start(); send.close()
            active.append((index, files[index], process, receive, time.monotonic()))
        for index, path, process, receive, worker_started in active:
            remaining = max(0.0, config.bounds.document_timeout_seconds - (time.monotonic() - worker_started))
            if receive.poll(remaining):
                try:
                    results[index] = receive.recv()
                except EOFError:
                    results[index] = None
            process.join(timeout=0.1)
            if process.is_alive():
                process.terminate(); process.join()
            if results[index] is None:
                document_id = _document_id(path)
                results[index] = RunResult(RunStatus.FAILED, f"failed:{document_id}", document_id, None, None, None, None, None, (f"document_worker_timeout_or_failure:{config.bounds.document_timeout_seconds}s",), time.monotonic() - worker_started)
            receive.close()
    return tuple(result for result in results if result is not None)


def _bounded_run_worker(send, input_path: Path, output: Path, config: RunConfig) -> None:
    try:
        _check_worker_memory(config.bounds)
        send.send(run_input(input_path, output, config))
    except BaseException as exc:
        document_id = _document_id(input_path)
        send.send(RunResult(RunStatus.FAILED, f"failed:{document_id}", document_id, None, None, None, None, None, (f"worker_failure:{type(exc).__name__}:{exc}",), 0.0))
    finally:
        send.close()


def _check_worker_memory(bounds: ResourceBounds) -> None:
    maximum = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    used_bytes = int(maximum if sys.platform == "darwin" else maximum * 1024)
    if used_bytes > bounds.worker_soft_memory_bytes:
        raise ResourceLimitError(f"worker_soft_memory_limit:{used_bytes}>{bounds.worker_soft_memory_bytes}")


def _cached_hypotheses(payload: object) -> tuple[HypothesisSet, ...]:
    if not isinstance(payload, list):
        raise ValueError("invalid mapper cache payload")
    result: list[HypothesisSet] = []
    for item in payload:
        if not isinstance(item, dict):
            raise ValueError("invalid mapper cache hypothesis")
        candidates = tuple(
            AssociationClaim(
                AssociationKind(candidate["kind"]),
                tuple(candidate["target_observation_ids"]),
                tuple(candidate.get("evidence_ids", ())),
                tuple(candidate.get("counterevidence_ids", ())),
                bool(candidate.get("source_sufficient", False)),
                candidate.get("semantic_field"),
                tuple(candidate.get("context_observation_ids", ())),
                tuple(tuple(pair) for pair in candidate.get("semantic_identity_assignments", ())),
                candidate.get("semantic_period_observation_id"),
                candidate.get("semantic_temporal_role"),
                candidate.get("semantic_unit_observation_id"),
                candidate.get("semantic_unit"),
                candidate.get("semantic_row_role"),
                tuple(candidate.get("header_parent_ids", ())),
                candidate.get("semantic_field_observation_id"),
                candidate.get("semantic_measure_observation_id"),
                candidate.get("semantic_measure"),
                candidate.get("semantic_movement_observation_id"),
                candidate.get("semantic_movement"),
                candidate.get("record_identity_row_id"),
                bool(candidate.get("period_conflict", False)),
            )
            for candidate in item["candidates"]
        )
        result.append(HypothesisSet(item["id"], item["document_id"], item["amount_observation_id"], candidates))
    return tuple(result)


def _verified_artifacts(directory: Path, manifest: dict[str, Any], expected_manifest_hash: str | None = None) -> bool:
    if manifest.get("readback_verified") is not True:
        return False
    files = manifest.get("files")
    if not isinstance(files, dict) or not files:
        return False
    for name, digest in files.items():
        path = directory / name
        if not path.is_file() or not isinstance(digest, str) or content_hash(path) != digest:
            return False
    manifest_path = directory / "manifest.v1.json"
    if expected_manifest_hash is not None and content_hash(manifest_path) != expected_manifest_hash:
        return False
    evidence = json.loads((directory / "evidence.v1.json").read_text())
    complete = (directory / ".complete").read_text().strip()
    document = evidence.get("document", {})
    return (
        complete == manifest.get("snapshot_digest") == evidence.get("snapshot", {}).get("digest")
        and manifest.get("status") == document.get("status")
        and manifest.get("run_id") == evidence.get("run_id")
        and manifest.get("document_id") == document.get("document_id")
        and manifest.get("source_hash") == document.get("source_hash")
        and manifest.get("snapshot_id") == evidence.get("snapshot", {}).get("id")
    )


def _observation_chunks(observations: tuple[SourceObservation, ...], bounds: ResourceBounds) -> tuple[tuple[SourceObservation, ...], ...]:
    pages = sorted({int(o.metadata["page_index"]) for o in observations if o.metadata.get("page_index") is not None})
    if pages:
        chunks: list[tuple[SourceObservation, ...]] = []
        step = bounds.pdf_chunk_pages - bounds.page_overlap
        if step <= 0:
            raise ValueError("pdf chunk overlap must be smaller than chunk size")
        for offset in range(0, len(pages), step):
            selected = set(pages[offset:offset + bounds.pdf_chunk_pages])
            if not selected:
                break
            chunks.append(tuple(o for o in observations if o.metadata.get("page_index") in selected))
            if offset + bounds.pdf_chunk_pages >= len(pages):
                break
        return tuple(chunks)
    by_sheet: dict[str, list[SourceObservation]] = {}
    for observation in observations:
        sheet = observation.locator.split("!", 1)[0]
        by_sheet.setdefault(sheet, []).append(observation)
    chunks = []
    for sheet_observations in by_sheet.values():
        coordinate_rows = {int(digits) for o in sheet_observations if isinstance(o.metadata.get("coordinate"), str) for digits in ("".join(character for character in o.metadata["coordinate"] if character.isdigit()),) if digits}
        rows = sorted({int(o.metadata["row"]) for o in sheet_observations if o.metadata.get("row") is not None} | coordinate_rows)
        if not rows:
            chunks.append(tuple(sheet_observations)); continue
        step = bounds.row_chunk_size - bounds.row_overlap
        if step <= 0:
            raise ValueError("row chunk overlap must be smaller than chunk size")
        # Context rows repeated in every window: the sheet's first rows (titles, reporting date) plus the header band of the
        # table the window starts in (the nearest band above its first row), never another table's band.
        blocks = [o for o in sheet_observations if o.kind == "record_block" and o.metadata.get("row") is not None]
        identity = {int(o.metadata["row"]) for o in blocks}
        members = {int(str(member).rsplit(":", 1)[1]) for o in blocks for member in o.metadata.get("member_ids", ()) if str(member).rsplit(":", 1)[1].isdigit()}
        bands = _header_bands(rows, identity, members, bounds.row_overlap)
        for offset in range(0, len(rows), step):
            window = set(rows[offset:offset + bounds.row_chunk_size]) | set(rows[:bounds.row_overlap]) | set(_band_for(bands, offset))
            def row_of(observation: SourceObservation) -> int | None:
                if observation.metadata.get("row") is not None:
                    return int(observation.metadata["row"])
                coordinate = observation.metadata.get("coordinate")
                if isinstance(coordinate, str):
                    digits = "".join(character for character in coordinate if character.isdigit())
                    return int(digits) if digits else None
                return None
            chunks.append(tuple(o for o in sheet_observations if o.kind == "table_region" or row_of(o) in window))
            if offset + bounds.row_chunk_size >= len(rows):
                break
    return tuple(chunks) if chunks else (observations,)


def _mapper_bounded_chunks(mapper: OpenAICompatibleMapper, document_id: str, observations: tuple[SourceObservation, ...], context_rows: int) -> tuple[tuple[SourceObservation, ...], ...]:
    if mapper.request_fits(document_id, observations):
        return (observations,)
    by_id = {observation.id: observation for observation in observations}
    regions = [observation for observation in observations if observation.kind == "table_region"]
    if not regions:
        raise MapperUnavailableError("mapper_minimum_coherent_region_exceeds_request_budget")

    def closure(region: SourceObservation, row_ids: list[str]) -> tuple[SourceObservation, ...]:
        selected_rows = {row_id for row_id in row_ids if row_id in by_id}
        # A record block travels whole: any selected member row pulls in the block observation and all of its rows,
        # identity row included, so the model can see and declare the record (fit checks run on this closure).
        blocks = [o for o in observations if o.kind == "record_block" and selected_rows.intersection(str(v) for v in o.metadata.get("member_ids", ()))]
        selected_rows.update(str(v) for block in blocks for v in block.metadata.get("member_ids", ()) if str(v) in by_id)
        selected_ids = set(selected_rows) | {block.id for block in blocks}
        for row_id in selected_rows:
            selected_ids.update(str(value) for value in by_id[row_id].metadata.get("member_ids", ()) if value in by_id)
        # Semantic header spans refer to the same source words as a selected
        # visual/logical row. Keep them with that complete row; never split a
        # record or manufacture a partial relationship graph.
        changed = True
        while changed:
            changed = False
            for observation in observations:
                if observation.kind != "logical_row" or observation.id in selected_ids:
                    continue
                members = {str(value) for value in observation.metadata.get("member_ids", ()) if value in by_id}
                if members and members.intersection(selected_ids):
                    selected_ids.add(observation.id); selected_ids.update(members); changed = True
        selected_ids.add(region.id)
        result: list[SourceObservation] = []
        for observation in observations:
            if observation.id not in selected_ids:
                continue
            if observation.id == region.id:
                ordered = list(dict.fromkeys((*(str(value) for value in region.metadata.get("member_ids", ())), *row_ids)))
                result.append(replace(region, metadata={**dict(region.metadata), "member_ids": tuple(row_id for row_id in ordered if row_id in selected_rows)}))
            else:
                result.append(observation)
        return tuple(result)

    pieces: list[tuple[SourceObservation, ...]] = []
    for region in regions:
        rows = [str(value) for value in region.metadata.get("member_ids", ()) if value in by_id]
        whole = closure(region, rows)
        if mapper.request_fits(document_id, whole):
            pieces.append(whole); continue
        blocks = [o for o in observations if o.kind == "record_block"]
        bands = _header_bands(rows, {str(o.metadata.get("identity_row_id")) for o in blocks}, {str(m) for o in blocks for m in o.metadata.get("member_ids", ())}, context_rows)
        top = rows[:context_rows]
        if top and not mapper.request_fits(document_id, closure(region, top)):
            raise MapperUnavailableError("mapper_minimum_coherent_region_exceeds_request_budget")
        pending = [row_id for row_id in rows if row_id not in set(top)]
        index_of = {row_id: index for index, row_id in enumerate(rows)}
        offset = 0
        while offset < len(pending):
            context = list(dict.fromkeys((*top, *_band_for(bands, index_of[pending[offset]]))))
            low, high, best = offset + 1, len(pending), None
            while low <= high:
                middle = (low + high) // 2
                candidate = [*context, *pending[offset:middle]]
                if mapper.request_fits(document_id, closure(region, candidate)):
                    best = middle; low = middle + 1
                else:
                    high = middle - 1
            if best is None:
                raise MapperUnavailableError("mapper_minimum_coherent_row_with_context_exceeds_request_budget")
            pieces.append(closure(region, [*context, *pending[offset:best]]))
            offset = best

    packed: list[tuple[SourceObservation, ...]] = []
    current: tuple[SourceObservation, ...] = ()
    current_regions: set[str] = set()
    for piece in pieces:
        piece_regions = {observation.id for observation in piece if observation.kind == "table_region"}
        combined_by_id = {observation.id: observation for observation in (*current, *piece)}
        combined = tuple(combined_by_id.values())
        can_combine = not current_regions.intersection(piece_regions) and mapper.request_fits(document_id, combined)
        if current and not can_combine:
            packed.append(current); current = piece; current_regions = piece_regions
        else:
            current = combined; current_regions.update(piece_regions)
    if current:
        packed.append(current)
    if not packed:
        raise MapperUnavailableError("mapper_no_coherent_regions")
    return tuple(packed)


def _regions_needing_mapper(observations: tuple[SourceObservation, ...], residual_ids: set[str]) -> tuple[SourceObservation, ...]:
    """Drop every table region whose amounts were all resolved locally (or that has none): the model cannot help there."""
    by_id = {o.id: o for o in observations}
    regions = [o for o in observations if o.kind == "table_region"]
    rows_of = {region.id: [str(row) for row in region.metadata.get("member_ids", ()) if str(row) in by_id] for region in regions}
    needed = {region.id for region in regions if any(str(member) in residual_ids for row in rows_of[region.id] for member in by_id[row].metadata.get("member_ids", ()))}
    if not needed:
        return ()
    keep_rows = {row for region in needed for row in rows_of[region]}
    drop_rows = {row for region in regions if region.id not in needed for row in rows_of[region.id]} - keep_rows
    keep_members = {str(member) for row in keep_rows for member in by_id[row].metadata.get("member_ids", ())}
    drop = {region.id for region in regions if region.id not in needed} | drop_rows | ({str(member) for row in drop_rows for member in by_id[row].metadata.get("member_ids", ())} - keep_members)
    return tuple(o for o in observations if o.id not in drop)


def _header_bands(rows: list, identity: set, members: set, limit: int) -> list[tuple[int, list]]:
    """(index of a table's first record, up to `limit` non-record rows directly above it) for every table on the sheet."""
    bands: list[tuple[int, list]] = []
    for index, row in enumerate(rows):
        if row not in identity or (index and rows[index - 1] in members):
            continue
        cursor = index; band: list = []
        while cursor and rows[cursor - 1] not in members and len(band) < limit:
            cursor -= 1; band.insert(0, rows[cursor])
        if band:
            bands.append((index, band))
    return bands


def _band_for(bands: list[tuple[int, list]], index: int) -> list:
    """The band of the table a window starting at `index` is in: the nearest band whose records start at or before it
    (a window still in the preamble gets the first table's band)."""
    above = [band for start, band in bands if start <= index]
    return above[-1] if above else (bands[0][1] if bands else [])


def _deduplicate_mapped(hypotheses: tuple[HypothesisSet, ...]) -> tuple[HypothesisSet, ...]:
    grouped: dict[str, dict[str, HypothesisSet]] = {}
    for hypothesis in hypotheses:
        # Overlapping requests yield the same claim under different plan digests: identical claims are one hypothesis;
        # claims that differ in any field or evidence still conflict and abstain.
        encoded = json.dumps([jsonable(candidate) for candidate in hypothesis.candidates], sort_keys=True)
        grouped.setdefault(hypothesis.amount_observation_id, {}).setdefault(encoded, hypothesis)
    return tuple(next(iter(variants.values())) for variants in grouped.values() if len(variants) == 1)


def _finalize_binding_id(binding: OutputBinding, snapshot_id: str, document_id: str) -> OutputBinding:
    payload = json.dumps({"snapshot_id": snapshot_id, "record_id": binding.record_id, "destination": jsonable(binding.destination), "status": binding.status, "transformations": dict(binding.transformations)}, sort_keys=True)
    binding_id = f"{document_id}:binding:{sha256(payload.encode()).hexdigest()[:20]}"
    return OutputBinding(binding_id, snapshot_id, binding.record_id, binding.destination, binding.source_pointer, binding.status, binding.transformations)
