from __future__ import annotations

from dataclasses import replace
from decimal import Decimal
from hashlib import sha256
import json
import re
import time

from .models import AssociationClaim, AssociationKind, Abstention, BusinessIdentity, CoverageDisposition, CoverageItem, HypothesisSet, PhysicalIdentity, SourceObservation, ValidatedFact
from .numeric import NumericParseError, parse_amount
from .policies import PolicyBundle
from .structure import ContextResolver, ResolvedContext, norm, unit_from_text


SEMANTICS_VERSION = "14"


def _stable_id(prefix: str, *parts: str) -> str:
    return f"{prefix}:{sha256('|'.join(parts).encode()).hexdigest()[:20]}"


class AssociationEngine:
    """Evidence gate: only unique row/column membership can promote a value."""

    def candidates(self, amount: SourceObservation, observations: tuple[SourceObservation, ...], context: ResolvedContext | None = None) -> HypothesisSet:
        same_doc = tuple(o for o in observations if o.document_id == amount.document_id and o.id != amount.id)
        claims: list[AssociationClaim] = []
        logical_rows = [o for o in same_doc if o.kind == "logical_row" and amount.id in o.metadata.get("member_ids", [])]
        visual_rows = [o for o in logical_rows if o.metadata.get("structure_role") == "visual_row"]
        if visual_rows:
            logical_rows = visual_rows
        for row in logical_rows:
            if context and context.business:
                claims.append(AssociationClaim(AssociationKind.OTHER, (row.id,), (row.id, *context.evidence_ids), (), True))
        claims.append(AssociationClaim(AssociationKind.UNKNOWN_UNMODELED, (), (), tuple(c.evidence_ids[0] for c in claims if c.evidence_ids), not claims))
        evidence_fingerprint = "|".join(context.evidence_ids if context else ())
        return HypothesisSet(_stable_id(f"{amount.document_id}:hyp", SEMANTICS_VERSION, amount.id, evidence_fingerprint), amount.document_id, amount.id, tuple(claims))

    def resolve(self, hypotheses: HypothesisSet) -> AssociationClaim | None:
        supported = [c for c in hypotheses.candidates if c.kind != AssociationKind.UNKNOWN_UNMODELED and c.source_sufficient and not c.counterevidence_ids]
        return supported[0] if len(supported) == 1 else None

    def validate_external(self, hypotheses: HypothesisSet, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> HypothesisSet:
        by_id = {o.id: o for o in observations if o.document_id == amount.document_id}
        validated: list[AssociationClaim] = []
        for candidate in hypotheses.candidates:
            if candidate.kind == AssociationKind.UNKNOWN_UNMODELED:
                validated.append(replace(candidate, source_sufficient=False))
                continue
            targets = [by_id.get(item) for item in candidate.target_observation_ids]
            relationship_ids = set(candidate.target_observation_ids + candidate.evidence_ids + candidate.context_observation_ids)
            regions = [
                observation for observation in by_id.values()
                if observation.id in relationship_ids
                and observation.kind in {"logical_row", "table_region"}
                and amount.id in observation.metadata.get("member_ids", ())
            ]
            # The model only names existing evidence. Sufficiency is established
            # locally by one record region containing the amount and every target
            # (a target may be a label/member rather than the region itself).
            explicit_membership = False
            if candidate.kind in {AssociationKind.GROUP, AssociationKind.OTHER} and len(regions) == 1 and targets and all(target is not None for target in targets):
                members = set(regions[0].metadata.get("member_ids", ()))
                explicit_membership = all(target.id == regions[0].id or target.id in members for target in targets if target is not None)
            validated.append(replace(candidate, source_sufficient=bool(explicit_membership)))
        if not any(c.kind == AssociationKind.UNKNOWN_UNMODELED for c in validated):
            validated.append(AssociationClaim(AssociationKind.UNKNOWN_UNMODELED, ()))
        return replace(hypotheses, candidates=tuple(validated))


def _numeric_like(text: str) -> bool:
    try:
        parse_amount(text)
        return True
    except NumericParseError:
        return False


def interpret(observations: tuple[SourceObservation, ...], coverage: tuple[CoverageItem, ...], policies: PolicyBundle, mapped_hypotheses: tuple[HypothesisSet, ...] = (), *, deadline: float | None = None) -> tuple[tuple[HypothesisSet, ...], tuple[ValidatedFact, ...], tuple[Abstention, ...], tuple[CoverageItem, ...]]:
    engine = AssociationEngine(); resolver = ContextResolver(observations)
    hypotheses: list[HypothesisSet] = []
    facts: list[ValidatedFact] = []
    abstentions: list[Abstention] = []
    by_locator: dict[str, str] = {}
    for obs in observations:
        if deadline is not None and time.monotonic() >= deadline:
            raise TimeoutError("document_timeout:during_semantics")
        is_explicit_empty = obs.kind == "empty_cell" and obs.channel == "xlsx_native" and obs.metadata.get("confirmed_empty") is True
        if obs.kind not in {"cell", "word", "empty_cell"}:
            continue
        if not is_explicit_empty and (not obs.text or not _numeric_like(obs.text)):
            continue
        if resolver.is_identity_value(obs, observations):
            continue
        context = resolver.resolve(obs, observations)
        mapped = [h for h in mapped_hypotheses if h.amount_observation_id == obs.id]
        local_hyp = engine.candidates(obs, observations, context)
        external_hyp = engine.validate_external(mapped[0], obs, observations) if len(mapped) == 1 else None
        mapped_dependencies: tuple[str, ...] = ()
        if external_hyp is not None:
            supported = [candidate for candidate in external_hyp.candidates if candidate.source_sufficient]
            if len(supported) == 1:
                context = _apply_mapped_axes(context, supported[0], obs, observations)
                mapped_dependencies = (external_hyp.id,)
                local_hyp = engine.candidates(obs, observations, context)
            hyp = _merge_hypotheses(local_hyp, external_hyp)
        else:
            hyp = local_hyp
        hypotheses.append(hyp)
        selected = engine.resolve(local_hyp) or (engine.resolve(external_hyp) if external_hyp is not None else None)
        if is_explicit_empty:
            accepted = policies.empty_zero.accepts(
                row_established=context.business is not None and context.business.row_kind == "data",
                amount_column_established=context.business is not None and context.business.field is not None,
                full_region_present=obs.metadata.get("full_region_present") is True,
                merged_follower=obs.metadata.get("merged_range") is not None and obs.metadata.get("merged_anchor") is not True,
                unreadable_or_omitted=False,
            )
            if not accepted:
                continue
            amount = replace(parse_amount("0", scale=context.scale), raw_lexeme="", sign_notation="implicit_zero")
        else:
            amount = parse_amount(obs.text or "", scale=context.scale)
        if selected is None or context.business is None or context.missing_axes:
            reason = "missing_context:" + ",".join(context.missing_axes) if context.missing_axes else "association_ambiguous_or_insufficient"
            abstention = Abstention(_stable_id(f"{obs.document_id}:abst", SEMANTICS_VERSION, obs.id, reason), obs.document_id, reason, (obs.id,), context.evidence_ids)
            abstentions.append(abstention); by_locator[obs.locator] = abstention.id
            continue
        label_id = selected.target_observation_ids[0]
        pointer = _pointer_for(obs)
        semantic_payload = json.dumps({
            "version": SEMANTICS_VERSION, "occurrence": obs.locator, "observation": obs.id,
            "association": {"hypothesis": hyp.id, "kind": selected.kind.value, "target": selected.target_observation_ids},
            "business": {name: str(getattr(context.business, name)) for name in context.business.__dataclass_fields__},
            "amount": {"source_signed": str(amount.source_signed_value), "normalized": str(amount.normalized_value), "scale": str(amount.scale), "precision": amount.precision},
            "policy": policies.digest(),
        }, sort_keys=True)
        fact_id = _stable_id(f"{obs.document_id}:fact", semantic_payload)
        dependencies = tuple(dict.fromkeys((obs.id, label_id, hyp.id, *mapped_dependencies, *context.evidence_ids)))
        fact = ValidatedFact(fact_id, obs.document_id, PhysicalIdentity(obs.source_hash, obs.document_id, obs.locator), context.business, amount, obs.id, pointer, dependencies, policies.versions())
        facts.append(fact); by_locator[obs.locator] = fact.id
    duplicate_groups: dict[tuple[object, ...], list[ValidatedFact]] = {}
    for fact in facts:
        business = fact.business
        row_ids = tuple(dependency for dependency in fact.dependency_ids if ":row:" in dependency or ":visual:" in dependency or ":ocrrow:" in dependency)
        key = (row_ids, business.entity, business.book, business.asset_id, business.asset_sub_id, business.class_code, business.balance_position, business.source_name, business.field, business.period, business.temporal_role, business.unit)
        duplicate_groups.setdefault(key, []).append(fact)
    duplicate_ids = {fact.id for group in duplicate_groups.values() if len(group) > 1 for fact in group}
    if duplicate_ids:
        retained: list[ValidatedFact] = []
        for fact in facts:
            if fact.id not in duplicate_ids:
                retained.append(fact)
                continue
            reason = "duplicate_same_row_field_ambiguous"
            abstention = Abstention(_stable_id(f"{fact.document_id}:abst", SEMANTICS_VERSION, fact.source_observation_id, reason), fact.document_id, reason, (fact.source_observation_id,), fact.dependency_ids)
            abstentions.append(abstention); by_locator[fact.physical.occurrence] = abstention.id
        facts = retained
    facts = list(_annotate_arithmetic(tuple(facts)))
    updated: list[CoverageItem] = []
    observations_by_locator = {o.locator: o for o in observations}
    dependency_to_fact = {dependency: fact.id for fact in facts for dependency in fact.dependency_ids}
    for item in coverage:
        match = by_locator.get(item.locator)
        if match:
            disposition = CoverageDisposition.FACT if any(f.id == match for f in facts) else CoverageDisposition.ABSTENTION
            updated.append(replace(item, disposition=disposition, linked_record_id=match, reason=None))
        elif item.kind in {"cell", "pdf_word"} and item.locator in observations_by_locator:
            observation = observations_by_locator[item.locator]
            linked_fact = dependency_to_fact.get(observation.id)
            if linked_fact:
                updated.append(replace(item, disposition=CoverageDisposition.FACT, linked_record_id=linked_fact, reason="semantic_evidence"))
            elif observation.kind == "cell" and observation.metadata.get("formula") and observation.metadata.get("cached_value") is None:
                updated.append(replace(item, disposition=CoverageDisposition.UNRESOLVED, reason="formula_result_unavailable"))
            elif observation.text and not _numeric_like(observation.text):
                updated.append(replace(item, disposition=CoverageDisposition.OUT_OF_SCOPE, reason="non_amount_structural_content"))
            elif resolver.is_identity_value(observation, observations):
                updated.append(replace(item, disposition=CoverageDisposition.OUT_OF_SCOPE, reason="business_identity_content"))
            else:
                updated.append(item)
        else:
            updated.append(item)
    return tuple(hypotheses), tuple(facts), tuple(abstentions), tuple(updated)


def _annotate_arithmetic(facts: tuple[ValidatedFact, ...]) -> tuple[ValidatedFact, ...]:
    grouped: dict[tuple[object, ...], list[ValidatedFact]] = {}
    for fact in facts:
        business = fact.business
        key = (business.entity, business.book, business.asset_id, business.asset_sub_id, business.class_code, business.balance_position, business.source_name, business.granularity, business.period)
        grouped.setdefault(key, []).append(fact)
    inconsistent: set[str] = set()
    required = {"ahk_open", "zugang", "abgang", "umb", "ahk_close"}
    for group in grouped.values():
        by_field: dict[str, list[ValidatedFact]] = {}
        for fact in group:
            if fact.business.field:
                by_field.setdefault(fact.business.field, []).append(fact)
        if not required <= set(by_field) or any(len(by_field[field]) != 1 for field in required):
            continue
        values = {field: by_field[field][0].amount.normalized_value for field in required}
        disposal = values["abgang"] if values["abgang"] < 0 else -values["abgang"]
        expected = values["ahk_open"] + values["zugang"] + disposal + values["umb"]
        if expected != values["ahk_close"]:
            inconsistent.update(fact.id for field in required for fact in by_field[field])
    return tuple(replace(fact, source_anomaly="arithmetic_inconsistency:ahk_rollforward") if fact.id in inconsistent else fact for fact in facts)


def _pointer_for(obs: SourceObservation) -> dict[str, object]:
    if obs.channel == "xlsx_native":
        sheet, cell = obs.locator.rsplit("!", 1)
        return {"kind": "xlsx", "sheet": sheet, "cell": cell, "merged_anchor": obs.metadata.get("merged_anchor"), "observation_id": obs.id}
    page_text = obs.locator.split(":", 2)[1]
    return {"kind": "pdf", "page_index": int(page_text) - 1, "bbox": list(obs.bbox or ()), "coordinate_space": "pymupdf_unrotated_points_top_left", "rotation": obs.metadata.get("page_rotation", 0), "observation_id": obs.id}


def _apply_mapped_axes(context: ResolvedContext, selected: AssociationClaim, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> ResolvedContext:
    if context.business is None:
        return context
    by_id = {o.id: o for o in observations}
    valid_headers = []
    if selected.semantic_field:
      field_evidence = (selected.semantic_field_observation_id,) if selected.semantic_field_observation_id else tuple(dict.fromkeys((*selected.context_observation_ids, *selected.evidence_ids)))
      for item in field_evidence:
        header = by_id.get(item)
        if header is None:
            continue
        if amount.channel == "xlsx_native" and header.channel == "xlsx_native":
            sheet, coordinate = amount.locator.rsplit("!", 1)
            if not header.locator.startswith(sheet + "!"):
                continue
            from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
            amount_letters, amount_row = coordinate_from_string(coordinate)
            letters, row = coordinate_from_string(header.locator.rsplit("!", 1)[1])
            if column_index_from_string(letters) == column_index_from_string(amount_letters) and row < amount_row:
                valid_headers.append(header)
        elif amount.channel in {"pdf_native", "pdf_ocr"} and header.bbox and amount.bbox:
            same_page = header.metadata.get("page_index") == amount.metadata.get("page_index")
            overlaps_x = min(header.bbox[2], amount.bbox[2]) > max(header.bbox[0], amount.bbox[0])
            if same_page and overlaps_x and header.bbox[3] < amount.bbox[1]:
                valid_headers.append(header)
    business = context.business
    evidence = list(context.evidence_ids)
    # A verified record block is a second opinion against local rules: disagreement abstains, it never overrides.
    block = selected.record_identity_row_id is not None
    conflicts: list[str] = []
    if block and selected.semantic_field and business.field not in {None, "amount", selected.semantic_field}:
        conflicts.append("field")
    if selected.semantic_field and len(valid_headers) == 1 and business.field in {None, "amount"}:
        business = replace(business, field=selected.semantic_field); evidence.append(valid_headers[0].id)
    target_rows = [by_id.get(value) for value in selected.target_observation_ids]
    target_rows = [row for row in target_rows if row is not None and amount.id in row.metadata.get("member_ids", ())]
    row_members = set(target_rows[0].metadata.get("member_ids", ())) if len(target_rows) == 1 else set()
    if block and len(target_rows) == 1 and by_id.get(selected.record_identity_row_id) is not None:
        row_members |= set(by_id[selected.record_identity_row_id].metadata.get("member_ids", ()))
    identity_updates: dict[str, str] = {}
    for observation_id, role in selected.semantic_identity_assignments:
        value = by_id.get(observation_id)
        if value is not None and observation_id in row_members and value.text and role in business.__dataclass_fields__ and getattr(business, role) is None:
            identity_updates[role] = value.text; evidence.append(observation_id)
        elif block and value is not None and role in business.__dataclass_fields__ and getattr(business, role) not in {None, value.text}:
            conflicts.append("business_identity")
    if identity_updates:
        business = replace(business, **identity_updates)
    # A model row label is descriptive context only. It cannot override the
    # deterministic record-role gate, especially for explicit empty cells.
    if selected.semantic_row_role and len(target_rows) == 1 and business.row_kind is None:
        business = replace(business, row_kind=selected.semantic_row_role); evidence.append(target_rows[0].id)
    unit = business.unit
    unit_observation = by_id.get(selected.semantic_unit_observation_id or "")
    if unit is None and unit_observation is not None and selected.semantic_unit in {"EUR", "TEUR"}:
        source_unit = unit_from_text(unit_observation.text)
        if source_unit == selected.semantic_unit:
            unit = source_unit; business = replace(business, unit=unit, currency="EUR"); evidence.append(unit_observation.id)
    period_observation = by_id.get(selected.semantic_period_observation_id or "")
    years = set(re.findall(r"\b(?:19|20)\d{2}\b", period_observation.text or "")) if period_observation is not None else set()
    # A verified block binds only under a period the mapper itself verified (R5/R2-6); a local header year alone is not enough.
    if block and (len(years) != 1 or (business.period is not None and business.period not in years)):
        conflicts.append("period")
    if business.period is None and len(years) == 1:
        business = replace(business, period=next(iter(years)), temporal_role=selected.semantic_temporal_role); evidence.append(period_observation.id)
    key = business.asset_id or business.class_code or business.balance_position or business.source_name
    business = replace(business, asset_or_class=key, granularity="asset" if business.asset_id else ("class" if business.class_code else "source_row"))
    recomputed = {
        "field": business.field not in {None, "amount"},
        "business_identity": bool(key),
        "unit": bool(unit),
        "period": business.period is not None,
    }
    # A sub-number clears only when the block bound one. Unresolved identifier columns clear only for an emitted block claim:
    # mapper R4 skips any block with a value under an undeclared identifier-like header anywhere above the records.
    carried = ({"asset_sub_id"} if "asset_sub_id" in identity_updates else set()) | {"unresolved_identifier"} if block and identity_updates else set()
    missing = tuple(dict.fromkeys((
        *(name for name in context.missing_axes if name not in recomputed and name not in carried),
        *(name for name, present in recomputed.items() if not present),
        *conflicts,
    )))
    return ResolvedContext(business, Decimal(1000) if unit == "TEUR" else Decimal(1), tuple(dict.fromkeys(evidence)), missing)


def _merge_hypotheses(local: HypothesisSet, external: HypothesisSet) -> HypothesisSet:
    candidates: list[AssociationClaim] = []
    seen: set[tuple[object, ...]] = set()
    for candidate in (*local.candidates, *external.candidates):
        key = tuple(str(value) for value in candidate.__dataclass_fields__) + (str(candidate),)
        if key not in seen:
            candidates.append(candidate); seen.add(key)
    digest = _stable_id(f"{local.document_id}:hyp", SEMANTICS_VERSION, local.id, external.id)
    return HypothesisSet(digest, local.document_id, local.amount_observation_id, tuple(candidates))


def prove_empty_zero(*, row_established: bool, amount_column_established: bool, full_region_present: bool, merged_follower: bool, unreadable_or_omitted: bool) -> bool:
    return row_established and amount_column_established and full_region_present and not merged_follower and not unreadable_or_omitted

