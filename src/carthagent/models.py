from __future__ import annotations

from dataclasses import dataclass, field, fields, is_dataclass
from datetime import date, datetime, time
from decimal import Decimal
from enum import StrEnum
from types import MappingProxyType
from typing import Any, Mapping


def _freeze(value: Any) -> Any:
    if isinstance(value, Mapping):
        return MappingProxyType({str(k): _freeze(v) for k, v in value.items()})
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(v) for v in value)
    if isinstance(value, (set, frozenset)):
        return frozenset(_freeze(v) for v in value)
    return value


def frozen_map(value: Mapping[str, Any] | None = None) -> Mapping[str, Any]:
    return _freeze(value or {})


class RunStatus(StrEnum):
    COMPLETE = "complete"
    PARTIAL = "partial"
    FAILED = "failed"
    UNSUPPORTED = "unsupported_input"


class CoverageDisposition(StrEnum):
    FACT = "linked_fact"
    ABSTENTION = "linked_abstention"
    OUT_OF_SCOPE = "evidenced_out_of_scope"
    UNRESOLVED = "unresolved_unsupported"


@dataclass(frozen=True, slots=True)
class PhysicalIdentity:
    source_hash: str
    document_id: str
    occurrence: str


@dataclass(frozen=True, slots=True)
class BusinessIdentity:
    entity: str | None = None
    book: str | None = None
    asset_or_class: str | None = None
    source_name: str | None = None
    asset_id: str | None = None
    asset_sub_id: str | None = None
    class_code: str | None = None
    class_name: str | None = None
    account: str | None = None
    balance_position: str | None = None
    granularity: str | None = None
    row_kind: str | None = None
    period: str | None = None
    temporal_role: str | None = None
    field: str | None = None
    currency: str | None = None
    unit: str | None = None


@dataclass(frozen=True, slots=True)
class OutputIdentity:
    contract_version: str
    row_key: str | None = None
    field: str | None = None
    cell: str | None = None


@dataclass(frozen=True, slots=True)
class SourceObservation:
    id: str
    document_id: str
    source_hash: str
    kind: str
    locator: str
    channel: str
    text: str | None = None
    bbox: tuple[float, float, float, float] | None = None
    metadata: Mapping[str, Any] = field(default_factory=frozen_map)

    def __post_init__(self) -> None:
        object.__setattr__(self, "metadata", frozen_map(self.metadata))
        if not self.id.startswith(f"{self.document_id}:"):
            raise ValueError("observation id must be document-local")


class AssociationKind(StrEnum):
    PREVIOUS = "previous"
    FOLLOWING = "following"
    GROUP = "group"
    OTHER = "other"
    UNKNOWN_UNMODELED = "unknown_unmodeled"


@dataclass(frozen=True, slots=True)
class AssociationClaim:
    kind: AssociationKind
    target_observation_ids: tuple[str, ...]
    evidence_ids: tuple[str, ...] = ()
    counterevidence_ids: tuple[str, ...] = ()
    source_sufficient: bool = False
    semantic_field: str | None = None
    context_observation_ids: tuple[str, ...] = ()
    semantic_identity_assignments: tuple[tuple[str, str], ...] = ()
    semantic_period_observation_id: str | None = None
    semantic_temporal_role: str | None = None
    semantic_unit_observation_id: str | None = None
    semantic_unit: str | None = None
    semantic_row_role: str | None = None
    header_parent_ids: tuple[str, ...] = ()
    semantic_field_observation_id: str | None = None
    semantic_measure_observation_id: str | None = None
    semantic_measure: str | None = None
    semantic_movement_observation_id: str | None = None
    semantic_movement: str | None = None
    record_identity_row_id: str | None = None
    # The plan's period contexts genuinely disagree on the year for this amount (item 7): the claim must abstain on period
    # and must not fall back to a local period either.
    period_conflict: bool = False


@dataclass(frozen=True, slots=True)
class HypothesisSet:
    id: str
    document_id: str
    amount_observation_id: str
    candidates: tuple[AssociationClaim, ...]

    def __post_init__(self) -> None:
        if not any(c.kind == AssociationKind.UNKNOWN_UNMODELED for c in self.candidates):
            raise ValueError("open-world hypothesis requires unknown_unmodeled")


@dataclass(frozen=True, slots=True)
class AmountRepresentation:
    raw_lexeme: str
    decimal_separator: str | None
    thousands_separator: str | None
    sign_notation: str
    displayed_value: Decimal
    source_signed_value: Decimal
    scale: Decimal
    normalized_value: Decimal
    precision: int


@dataclass(frozen=True, slots=True)
class ValidatedFact:
    id: str
    document_id: str
    physical: PhysicalIdentity
    business: BusinessIdentity
    amount: AmountRepresentation
    source_observation_id: str
    source_pointer: Mapping[str, Any]
    dependency_ids: tuple[str, ...]
    policy_versions: Mapping[str, str]
    source_anomaly: str = "none"

    def __post_init__(self) -> None:
        object.__setattr__(self, "policy_versions", frozen_map(self.policy_versions))
        object.__setattr__(self, "source_pointer", frozen_map(self.source_pointer))


@dataclass(frozen=True, slots=True)
class Abstention:
    id: str
    document_id: str
    reason: str
    candidate_observation_ids: tuple[str, ...]
    dependency_ids: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class CoverageItem:
    id: str
    document_id: str
    kind: str
    locator: str
    disposition: CoverageDisposition
    linked_record_id: str | None = None
    reason: str | None = None


@dataclass(frozen=True, slots=True)
class OutputBinding:
    id: str
    snapshot_id: str
    record_id: str
    destination: OutputIdentity
    source_pointer: Mapping[str, Any] = field(default_factory=frozen_map)
    status: str = "contract_pending"
    transformations: Mapping[str, Any] = field(default_factory=frozen_map)

    def __post_init__(self) -> None:
        object.__setattr__(self, "source_pointer", frozen_map(self.source_pointer))
        object.__setattr__(self, "transformations", frozen_map(self.transformations))


def jsonable(record: Any) -> dict[str, Any]:
    def convert(value: Any) -> Any:
        if is_dataclass(value) and not isinstance(value, type):
            return {item.name: convert(getattr(value, item.name)) for item in fields(value)}
        if isinstance(value, Decimal):
            return str(value)
        if isinstance(value, (datetime, date, time)):
            return value.isoformat()
        if isinstance(value, Mapping):
            return {k: convert(v) for k, v in value.items()}
        if isinstance(value, (tuple, list, set, frozenset)):
            return [convert(v) for v in value]
        if isinstance(value, StrEnum):
            return value.value
        return value
    converted = convert(record)
    if not isinstance(converted, dict):
        raise TypeError("record must serialize to an object")
    return converted

