from __future__ import annotations

from dataclasses import asdict, dataclass, field, fields
from hashlib import sha256
import json


@dataclass(frozen=True, slots=True)
class ResourceBounds:
    max_input_bytes: int = 200 * 1024 * 1024
    max_pdf_pages: int = 500
    max_worksheets: int = 100
    max_cells: int = 2_000_000
    max_page_megapixels: int = 40
    ocr_dpi: int = 300
    document_timeout_seconds: int = 600
    worker_processes: int = 2
    transient_retries: int = 2
    pdf_chunk_pages: int = 20
    row_chunk_size: int = 250
    page_overlap: int = 2
    row_overlap: int = 20
    worker_soft_memory_bytes: int = 2 * 1024 * 1024 * 1024
    mapper_phase_timeout_seconds: int = 480
    post_mapper_reserve_seconds: int = 120
    max_mapper_requests: int = 6
    max_mapper_response_bytes: int = 8_388_608

    def __post_init__(self) -> None:
        if isinstance(self.max_mapper_response_bytes, bool) or not isinstance(self.max_mapper_response_bytes, int) or self.max_mapper_response_bytes <= 0:
            raise ValueError("max_mapper_response_bytes must be a positive integer")


@dataclass(frozen=True, slots=True)
class VersionedPolicy:
    name: str
    version: str
    contract_pending: bool = True


@dataclass(frozen=True, slots=True)
class PolicyBundle:
    input_scope: VersionedPolicy = field(default_factory=lambda: InputScopePolicy("InputScopePolicy", "provisional-1"))
    excel_value: VersionedPolicy = field(default_factory=lambda: ExcelValuePolicy("ExcelValuePolicy", "provisional-1"))
    empty_zero: VersionedPolicy = field(default_factory=lambda: EmptyZeroPolicy("EmptyZeroPolicy", "provisional-1"))
    sign_polarity: VersionedPolicy = field(default_factory=lambda: SignPolarityPolicy("SignPolarityPolicy", "provisional-1"))
    pointer: VersionedPolicy = field(default_factory=lambda: PointerPolicy("PointerPolicy", "provisional-1"))
    destination_identity: VersionedPolicy = field(default_factory=lambda: DestinationIdentityPolicy("DestinationIdentityPolicy", "pending-1"))
    output_contract: VersionedPolicy = field(default_factory=lambda: OutputContract("OutputContract", "provisional-evidence-v1"))
    export_adapter: VersionedPolicy = field(default_factory=lambda: ExportAdapter("ExportAdapter", "provisional-review-v1"))
    scorer_adapter: VersionedPolicy = field(default_factory=lambda: ScorerAdapter("ScorerAdapter", "pending-1"))

    def versions(self) -> dict[str, str]:
        return {policy.name: policy.version for item in fields(self) for policy in (getattr(self, item.name),)}

    def digest(self) -> str:
        payload = json.dumps(asdict(self), sort_keys=True, separators=(",", ":"))
        return sha256(payload.encode()).hexdigest()


class InputScopePolicy(VersionedPolicy):
    def classify(self, suffix: str) -> str:
        suffix = suffix.lower()
        return "supported" if suffix in {".pdf", ".xlsx"} else ("unsupported_input" if suffix in {".doc", ".docx"} else "out_of_scope")


class ExcelValuePolicy(VersionedPolicy):
    def select(self, *, raw: object, formula: str | None, cached: object, display: object) -> tuple[str, object | None]:
        if formula and cached is None: return "formula_result_unavailable", None
        if formula and display is not None and str(display) != str(cached): return "representation_conflict", None
        return ("cached_formula_result", cached) if formula else ("native_raw", raw)


class EmptyZeroPolicy(VersionedPolicy):
    def accepts(self, *, row_established: bool, amount_column_established: bool, full_region_present: bool, merged_follower: bool, unreadable_or_omitted: bool) -> bool:
        return row_established and amount_column_established and full_region_present and not merged_follower and not unreadable_or_omitted


class SignPolarityPolicy(VersionedPolicy):
    def target(self, source_signed: int | float, convention: str | None) -> int | float | None:
        if convention is None: return None
        if convention == "source_signed": return source_signed
        if convention == "positive_magnitude": return abs(source_signed)
        raise ValueError("unknown sign convention")


class PointerPolicy(VersionedPolicy):
    def validate(self, pointer: dict[str, object]) -> bool:
        if pointer.get("kind") == "xlsx": return bool(pointer.get("sheet") and pointer.get("cell") and pointer.get("observation_id"))
        box = pointer.get("bbox")
        return pointer.get("kind") == "pdf" and isinstance(box, (list, tuple)) and len(box) == 4 and bool(pointer.get("observation_id"))


class DestinationIdentityPolicy(VersionedPolicy):
    def status(self, row_key: str | None, field: str | None) -> str:
        return "bound" if row_key and field and not self.contract_pending else "contract_pending"


class OutputContract(VersionedPolicy):
    def permits(self, binding_status: str) -> bool: return not self.contract_pending and binding_status == "bound"


class ExportAdapter(VersionedPolicy):
    def artifact_kind(self) -> str: return "provisional_review" if self.contract_pending else "official"


class ScorerAdapter(VersionedPolicy):
    def available(self) -> bool: return not self.contract_pending


DEFAULT_BOUNDS = ResourceBounds()
DEFAULT_POLICIES = PolicyBundle()

