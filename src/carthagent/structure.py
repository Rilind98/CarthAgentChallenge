from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache
import re
import unicodedata

from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, range_boundaries

from .models import BusinessIdentity, SourceObservation


FIELD_ALIASES = {
 "ahk_open": ("ahk beginn", "ahk 01 01", "anschaffungskosten beginn", "opening cost"), "zugang": ("zugang", "zugänge", "additions"),
 "abgang": ("abgang", "abgänge", "disposals"), "umb": ("umbuchung", "umbuchungen", "transfers"),
 "ahk_close": ("ahk ende", "ahk 31 12", "anschaffungskosten ende", "closing cost"), "kum_afa_open": ("kum afa beginn", "opening accumulated depreciation"),
 "afa_plan": ("planmäßige afa", "planned depreciation"), "total_afa_gj": ("afa gj", "geschäftsjahr", "depreciation current year", "total depreciation"),
 "extraordinary": ("außerplanmäßige afa", "extraordinary depreciation"), "afa_abgang": ("afa abgang", "depreciation on disposals"),
 "zuschreibung": ("zuschreibung", "zuschreib", "write up"), "afa_umb": ("afa umbuchung", "depreciation transfer"),
 "kum_afa_close": ("kum afa ende", "closing accumulated depreciation"), "bw_close": ("buchwert ende", "buchwert 31 12", "closing carrying amount", "nbv closing"),
 "bw_prev": ("buchwert vorjahr", "prior year carrying amount"), "fk_zinsen": ("fremdkapitalzinsen", "borrowing costs"),
 "bew_ahk": ("bewertung ahk",), "bew_wb": ("bewertung wiederbeschaffung", "wiederbeschaffungswert"),
 "abgang_bw": ("buchwert abgang", "abgang buchwert"), "amount": ("betrag", "amount")}
FIELDS = FIELD_ALIASES
IDENTITIES = {"asset_id": ("anlagennummer", "anlagen nr", "anl nr", "inventarnummer", "inventar nummer", "inventar nr", "asset id", "asset number", "asset no"), "asset_sub_id": ("unternummer", "unter nr", "unternr", "subnumber", "sub id"), "source_name": ("bezeichnung", "asset name", "description", "name"), "class_code": ("anlagenklasse", "class code", "asset class no"), "class_name": ("anlagenklassenbezeichnung", "class name", "asset class name"), "account": ("konto", "account"), "balance_position": ("bilanzposition", "balance position"), "entity": ("gesellschaft", "company code", "entity"), "book": ("ledger", "accounting book", "bewertungsbereich")}
HEADER_GROUPS = {
    "ahk": ("anschaffungs und herstellungskosten", "anschaffungskosten", "ansch herst k", "ahk", "historical cost"),
    "cum_depreciation": ("kumulierte abschreibungen", "kumulierte afa", "abschreibung", "absch", "afa", "accumulated depreciation"),
    "book_value": ("buchwerte", "buchwert", "bw", "carrying amounts", "net book values"),
}


_MOVEMENT_WORD = re.compile(r"zug[aä]ng|abg[aä]ng|umbuch|zuschreib")
_DEPRECIATION_WORD = re.compile(r"\bafa\b|abschreib")
_PERIOD_QUALIFIER = re.compile(r"\b(?:im|des|geschäftsjahr(?:es)?)\b")


@lru_cache(maxsize=32_768)
def norm(text: str) -> str:
    return " ".join(re.sub(r"[^\wäöüß]+", " ", unicodedata.normalize("NFKC", text).casefold()).split())


@lru_cache(maxsize=32_768)
def _match_known(text: str, vocabulary_name: str) -> str | None:
    vocabulary = {"fields": FIELDS, "identities": IDENTITIES, "groups": HEADER_GROUPS}[vocabulary_name]
    return _match_uncached(text, vocabulary)


def _match_uncached(text: str, vocab: dict[str, tuple[str, ...]]) -> str | None:
    value = norm(text)
    scores: dict[str, int] = {}
    for key, terms in vocab.items():
        matched = [norm(term) for term in terms if norm(term) == value or norm(term) in value]
        if matched:
            scores[key] = max(len(term) + (10_000 if term == value else 0) for term in matched)
    if not scores:
        return None
    best = max(scores.values())
    winners = [key for key, score in scores.items() if score == best]
    return winners[0] if len(winners) == 1 else None


def match(text: str, vocab: dict[str, tuple[str, ...]]) -> str | None:
    if vocab is FIELDS:
        value = norm(text)
        if "abgang" in value and "minderung" in value:
            return None
        # "Zugänge im Geschäftsjahr" names additions; the period qualifier only
        # means current-year depreciation when no movement word is present.
        if "geschäftsjahr" in value and _MOVEMENT_WORD.search(value) and not _DEPRECIATION_WORD.search(value):
            return _match_uncached(_PERIOD_QUALIFIER.sub(" ", value), FIELDS)
        return _match_known(text, "fields")
    if vocab is IDENTITIES:
        return _match_known(text, "identities")
    if vocab is HEADER_GROUPS:
        return _match_known(text, "groups")
    return _match_uncached(text, vocab)


@dataclass(frozen=True, slots=True)
class ResolvedContext:
    business: BusinessIdentity | None
    scale: Decimal
    evidence_ids: tuple[str, ...]
    missing_axes: tuple[str, ...]


class ContextResolver:
    def __init__(self, observations: tuple[SourceObservation, ...] = ()) -> None:
        self._observations = observations
        self._xlsx_indexes: dict[tuple[str, str], list[tuple[int, int, SourceObservation]]] = {}
        self._xlsx_by_row: dict[tuple[str, str], dict[int, list[tuple[int, SourceObservation]]]] = {}
        self._xlsx_by_col: dict[tuple[str, str], dict[int, list[tuple[int, SourceObservation]]]] = {}
        self._xlsx_header_scopes: dict[tuple[str, str, int], tuple[int, int]] = {}
        self._xlsx_record_roles: dict[tuple[str, str, int], str] = {}
        self._xlsx_axes: dict[tuple[str, str, int, int, int], tuple[str | None, str | None, str | None, tuple[str, ...]]] = {}
        self._xlsx_period_sets: dict[tuple[str, str, int, int], tuple[set[str], set[str]]] = {}
        self._pdf_rows: dict[str, str] = {}
        self._observations_by_id = {observation.id: observation for observation in observations}
        self._recognition_candidates_by_anchor: dict[str, list[SourceObservation]] = {}
        self._pdf_record_block_by_row: dict[str, SourceObservation] = {}
        self._pdf_layouts: dict[tuple[str, str, str], _PdfLayout | None] = {}
        for observation in observations:
            if observation.kind == "logical_row" and observation.metadata.get("structure_role") == "visual_row":
                for member_id in observation.metadata.get("member_ids", ()):
                    self._pdf_rows[str(member_id)] = observation.id
            elif observation.kind == "record_block":
                for row_id in observation.metadata.get("member_ids", ()):
                    self._pdf_record_block_by_row[str(row_id)] = observation
            elif observation.kind == "recognition_candidate" and observation.metadata.get("anchor_observation_id"):
                self._recognition_candidates_by_anchor.setdefault(str(observation.metadata["anchor_observation_id"]), []).append(observation)

    def resolve(self, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> ResolvedContext:
        return self._xlsx(amount, observations) if amount.channel == "xlsx_native" else self._pdf(amount, observations)

    def _xlsx(self, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> ResolvedContext:
        sheet, coord = amount.locator.rsplit("!", 1); letters, row = coordinate_from_string(coord); col = column_index_from_string(letters)
        cache_key = (amount.document_id, sheet)
        indexed = self._xlsx_index(amount.document_id, sheet, observations)
        by_row = self._xlsx_by_row[cache_key]; by_col = self._xlsx_by_col[cache_key]
        scope_key = (amount.document_id, sheet, row)
        if scope_key not in self._xlsx_header_scopes:
            self._xlsx_header_scopes[scope_key] = _xlsx_header_scope(row, indexed)
        header_start, header_cutoff = self._xlsx_header_scopes[scope_key]
        axis_key = (amount.document_id, sheet, col, header_start, header_cutoff)
        if axis_key not in self._xlsx_axes:
            self._xlsx_axes[axis_key] = _xlsx_axis(col, row, indexed, header_start, header_cutoff)
        field, period, temporal_role, axis_evidence = self._xlsx_axes[axis_key]
        row_measure_observations = [
            observation for _, observation in by_row.get(row, ())
            if observation.text and match(observation.text, HEADER_GROUPS)
        ]
        row_measure = (
            match(row_measure_observations[0].text or "", HEADER_GROUPS)
            if len(row_measure_observations) == 1 else None
        )
        if row_measure is not None:
            field = compose_matrix_axes(row_measure, field, temporal_role)
            axis_evidence = tuple(dict.fromkeys((*axis_evidence, row_measure_observations[0].id)))
        period_key = (amount.document_id, sheet, header_start, header_cutoff)
        if period_key not in self._xlsx_period_sets:
            axis_periods = []
            for candidate in sorted({c for c, _, _ in indexed}):
                candidate_key = (amount.document_id, sheet, candidate, header_start, header_cutoff)
                if candidate_key not in self._xlsx_axes:
                    self._xlsx_axes[candidate_key] = _xlsx_axis(candidate, row, indexed, header_start, header_cutoff)
                axis_periods.append(self._xlsx_axes[candidate_key][:2])
            periods_by_field: dict[str, set[str]] = {}
            for candidate_field, candidate_period in axis_periods:
                if candidate_field and candidate_period:
                    periods_by_field.setdefault(candidate_field, set()).add(candidate_period)
            self._xlsx_period_sets[period_key] = ({value for values in periods_by_field.values() for value in values}, {name for name, values in periods_by_field.items() if len(values) > 1})
        all_periods, parallel_fields = self._xlsx_period_sets[period_key]
        sheet_header_periods = {
            year for _, header_row, observation in indexed if header_start <= header_row <= header_cutoff
            for year in _reported_periods(observation.text or "")
        }
        reporting_periods = all_periods | sheet_header_periods
        if reporting_periods and field not in parallel_fields:
            period = max(reporting_periods)
        local_unit_evidence = [o for header_row, o in by_col.get(col, ()) if o.text and header_start <= header_row <= header_cutoff and _header_span(col, o) and re.search(r"\b(?:T?EUR|T€)\b", o.text, re.I)]
        currency_columns = {candidate_col for candidate_col, candidate_row, observation in indexed if header_start <= candidate_row <= header_cutoff and re.search(r"(?:währ\s*ung|waehr\s*ung|currency)", norm(observation.text or ""), re.I)}
        record_unit_evidence = [observation for candidate_col, observation in by_row.get(row, ()) if candidate_col in currency_columns and re.fullmatch(r"\s*(?:T?EUR|T€)\s*", observation.text or "", re.I)]
        global_unit_evidence = [o for _, r, o in indexed if header_start <= r <= header_cutoff and re.search(r"\b(?:T?EUR|T€)\b", o.text or "", re.I) and (_spans_multiple_columns(o) or sum(bool(candidate.text) for _, candidate in by_row.get(r, ())) == 1)]
        unit_evidence = local_unit_evidence or record_unit_evidence or global_unit_evidence
        units = {"TEUR" if re.search(r"\b(?:TEUR|T€)\b", o.text or "", re.I) else "EUR" for o in unit_evidence}
        unit = next(iter(units)) if len(units) == 1 else None
        identity: dict[str, str] = {}; evidence = [*axis_evidence, *(o.id for o in unit_evidence)]
        for c, value in by_row.get(row, ()):
            if c == col: continue
            identity_headers = [(hr, hit) for hr, o in by_col.get(c, ()) if header_start <= hr <= header_cutoff and (hit := match(o.text or "", IDENTITIES))]
            closest_header_row = max((hr for hr, _ in identity_headers), default=None)
            axes = {hit for hr, hit in identity_headers if hr == closest_header_row}
            if len(axes) == 1:
                identity_value = value
                if not identity_value.text and isinstance(identity_value.metadata.get("merged_range"), str):
                    merged_range = identity_value.metadata["merged_range"]
                    anchors = [
                        candidate for _, _, candidate in indexed
                        if candidate.metadata.get("merged_range") == merged_range
                        and candidate.metadata.get("merged_anchor") is True
                        and candidate.text
                    ]
                    identity_value = anchors[0] if len(anchors) == 1 else identity_value
                if identity_value.text:
                    axis = next(iter(axes)); identity[axis] = identity_value.text; evidence.append(identity_value.id)
        unresolved_identifier_columns = {
            candidate_col for candidate_col, header_row, header in indexed
            if header_start <= header_row <= header_cutoff
            and match(header.text or "", IDENTITIES) is None
            and re.search(r"\b(?:nummer|number|code|identifier|record id)\b", norm(header.text or ""))
        }
        unresolved_identifier_present = any(candidate_col in unresolved_identifier_columns and value.text for candidate_col, value in by_row.get(row, ()))
        declared_sub_id = any(match(header.text or "", IDENTITIES) == "asset_sub_id" and header_start <= header_row <= header_cutoff for _, header_row, header in indexed)
        role_key = (amount.document_id, sheet, row)
        if role_key not in self._xlsx_record_roles:
            self._xlsx_record_roles[role_key] = _xlsx_record_role(row, indexed)
        role = self._xlsx_record_roles[role_key]
        key = identity.get("asset_id") or identity.get("class_code") or identity.get("balance_position") or (None if unresolved_identifier_present else identity.get("source_name"))
        missing = tuple(name for name, ok in (("field", field not in {None, "amount"}), ("business_identity", key), ("asset_sub_id", not declared_sub_id or identity.get("asset_sub_id") is not None), ("period", period is not None), ("unit", unit), ("unresolved_identifier", not unresolved_identifier_present)) if not ok)
        business = BusinessIdentity(entity=identity.get("entity"), book=identity.get("book"), asset_or_class=key, source_name=identity.get("source_name"), asset_id=identity.get("asset_id"), asset_sub_id=identity.get("asset_sub_id"), class_code=identity.get("class_code"), class_name=identity.get("class_name"), account=identity.get("account"), balance_position=identity.get("balance_position"), granularity="asset" if identity.get("asset_id") else ("class" if identity.get("class_code") else "source_row"), row_kind=role, period=period, temporal_role=temporal_role, field=field, currency="EUR", unit=unit)
        return ResolvedContext(business, Decimal(1000) if unit == "TEUR" else Decimal(1), tuple(dict.fromkeys(evidence)), missing)

    def is_identity_value(self, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> bool:
        if amount.channel in {"pdf_native", "pdf_ocr"}:
            if not amount.bbox:
                return False
            layout = self._pdf_layout(amount, observations)
            return bool(layout and layout.axes and (amount.bbox[0] + amount.bbox[2]) / 2 < layout.axes[0].left)
        if amount.channel != "xlsx_native": return False
        sheet, coord = amount.locator.rsplit("!", 1); letters, row = coordinate_from_string(coord); col = column_index_from_string(letters)
        self._xlsx_index(amount.document_id, sheet, observations)
        return any(header_row < row and match(header.text or "", IDENTITIES) for header_row, header in self._xlsx_by_col[(amount.document_id, sheet)].get(col, ()))

    def _pdf(self, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> ResolvedContext:
        if not amount.bbox: return ResolvedContext(None, Decimal(1), (), ("pointer",))
        layout = self._pdf_layout(amount, observations)
        if layout is None:
            return ResolvedContext(None, Decimal(1), (), ("logical_row", "column_structure"))
        amount_center = (amount.bbox[0] + amount.bbox[2]) / 2
        selected = next((axis for axis in layout.axes if axis.left <= amount_center <= axis.right), None)
        member_ids = set(layout.row.metadata.get("member_ids", ()))
        block = self._pdf_record_block_by_row.get(layout.row.id)
        if block is not None:
            identity_row = self._observations_by_id.get(str(block.metadata.get("identity_row_id")))
            if identity_row is not None:
                member_ids.update(str(value) for value in identity_row.metadata.get("member_ids", ()))
        identity_owner_row_ids = {layout.row.id}
        if block is not None and block.metadata.get("identity_row_id"):
            identity_owner_row_ids.add(str(block.metadata["identity_row_id"]))
        identity: dict[str, list[str]] = {}
        recognition_evidence_ids: list[str] = []
        for word in layout.words:
            if word.id not in member_ids or not word.bbox or not layout.axes:
                continue
            center = (word.bbox[0] + word.bbox[2]) / 2
            for identity_axis in layout.identity_axes:
                if identity_axis.left <= center < identity_axis.right:
                    value = word.text or ""
                    candidates = self._recognition_candidates_by_anchor.get(word.id, ())
                    usable_candidates = [
                        candidate for candidate in candidates
                        if candidate.text
                        and candidate.metadata.get("ownership_closed") is True
                        and "table" in candidate.metadata.get("provider_scopes", ())
                        and candidate.metadata.get("authoritative_row_id") in identity_owner_row_ids
                        and candidate.metadata.get("left_anchor_observation_id")
                        and candidate.metadata.get("right_anchor_observation_id")
                    ]
                    if identity_axis.role in {"asset_id", "asset_sub_id", "class_code", "account"} and len(usable_candidates) == 1:
                        value = usable_candidates[0].text or value
                        recognition_evidence_ids.append(usable_candidates[0].id)
                    identity.setdefault(identity_axis.role, []).append(value)
                    break
        identity_values = {axis: " ".join(parts).strip() for axis, parts in identity.items()}
        asset_id = identity_values.get("asset_id")
        source_name = identity_values.get("source_name")
        identity_disagrees = bool(
            amount.channel == "pdf_ocr"
            and (
                (asset_id and not _pdf_ocr_identity_consistent(layout, asset_id))
                or _ocr_identity_text_has_noise(asset_id)
                or _ocr_identity_text_has_noise(source_name)
            )
        )
        if identity_disagrees:
            asset_id = None
        unresolved_leading_identifier = bool(source_name and re.match(r"^\s*\d+\b", source_name) and not asset_id and not identity_values.get("class_code"))
        business_key = None if identity_disagrees or unresolved_leading_identifier else (asset_id or identity_values.get("class_code") or identity_values.get("balance_position") or source_name)
        row_text = norm(layout.row.text or "")
        is_page_footer = bool(
            re.search(r"\b(?:seite|page)\s*\d+\b", row_text)
            and re.search(r"\b(?:datum|date|printed|erstellt)\b", row_text)
        )
        row_kind = "note" if is_page_footer else ("subtotal" if row_text.startswith(("summe ", "subtotal ")) else ("total" if row_text.startswith(("gesamt ", "total ")) else "data"))
        row_member_ids = set(layout.row.metadata.get("member_ids", ()))
        measure_spans = [
            observation for observation in observations
            if observation.metadata.get("semantic_group")
            and row_member_ids.intersection(str(value) for value in observation.metadata.get("member_ids", ()))
        ]
        measure = str(measure_spans[0].metadata.get("semantic_group")) if len(measure_spans) == 1 else None
        field = _compose_matrix_field(measure, selected, amount) if selected else None
        period = selected.period if selected and (selected.parallel_period or layout.reporting_period is None) else layout.reporting_period
        unit, unit_evidence = _pdf_unit(amount, layout, selected)
        declared_sub_id = any(axis.role == "asset_sub_id" for axis in layout.identity_axes)
        missing = tuple(name for name, ok in (("field", field not in {None, "amount"}), ("business_identity", business_key), ("asset_sub_id", not declared_sub_id or identity_values.get("asset_sub_id") is not None), ("period", period is not None), ("unit", unit), ("logical_row", layout.row is not None), ("record_block", measure is None or block is not None), ("column_structure", selected is not None), ("unresolved_identifier", not unresolved_leading_identifier), ("ocr_identity_disagreement", not identity_disagrees), ("non_data_note", row_kind != "note")) if not ok)
        evidence = tuple(dict.fromkeys((*(selected.evidence_ids if selected else ()), *(axis.evidence_id for axis in layout.identity_axes), *recognition_evidence_ids, *(span.id for span in measure_spans), *((block.id,) if block else ()), *unit_evidence, *(word.id for word in layout.words if word.id in member_ids), layout.row.id)))
        business = BusinessIdentity(asset_or_class=business_key, source_name=source_name, asset_id=asset_id, asset_sub_id=identity_values.get("asset_sub_id"), class_code=identity_values.get("class_code"), class_name=identity_values.get("class_name"), account=identity_values.get("account"), balance_position=identity_values.get("balance_position"), entity=identity_values.get("entity"), book=identity_values.get("book"), granularity="asset" if asset_id else ("class" if identity_values.get("class_code") else "source_row"), row_kind=row_kind, period=period, temporal_role=selected.temporal_role if selected else None, field=field, currency="EUR", unit=unit)
        return ResolvedContext(business, Decimal(1000) if unit == "TEUR" else Decimal(1), evidence, missing)

    def _pdf_layout(self, amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> _PdfLayout | None:
        page = amount.locator.split(":", 2)[1]
        row_id = self._pdf_rows.get(amount.id, amount.id)
        key = (amount.document_id, page, row_id)
        if key not in self._pdf_layouts:
            self._pdf_layouts[key] = _pdf_layout(amount, observations)
        return self._pdf_layouts[key]

    def _xlsx_index(self, document_id: str, sheet: str, observations: tuple[SourceObservation, ...]) -> list[tuple[int, int, SourceObservation]]:
        key = (document_id, sheet)
        if key in self._xlsx_indexes:
            return self._xlsx_indexes[key]
        indexed: list[tuple[int, int, SourceObservation]] = []
        by_row: dict[int, list[tuple[int, SourceObservation]]] = {}
        by_col: dict[int, list[tuple[int, SourceObservation]]] = {}
        for observation in observations:
            if observation.channel != "xlsx_native" or not observation.locator.startswith(sheet + "!"):
                continue
            letters, row = coordinate_from_string(observation.locator.rsplit("!", 1)[1]); column = column_index_from_string(letters)
            indexed.append((column, row, observation)); by_row.setdefault(row, []).append((column, observation)); by_col.setdefault(column, []).append((row, observation))
        self._xlsx_indexes[key] = indexed; self._xlsx_by_row[key] = by_row; self._xlsx_by_col[key] = by_col
        return indexed


@dataclass(frozen=True, slots=True)
class _PdfIdentityAxis:
    left: float
    right: float
    role: str
    evidence_id: str


@dataclass(frozen=True, slots=True)
class _PdfAxis:
    left: float
    right: float
    field: str | None
    period: str | None
    temporal_role: str | None
    group: str | None
    header_text: str
    evidence_ids: tuple[str, ...]
    parallel_period: bool = False


@dataclass(frozen=True, slots=True)
class _PdfLayout:
    row: SourceObservation
    words: tuple[SourceObservation, ...]
    page_rows: tuple[SourceObservation, ...]
    axes: tuple[_PdfAxis, ...]
    identity_axes: tuple[_PdfIdentityAxis, ...]
    reporting_period: str | None
    header_top: float
    header_bottom: float


def _pdf_span_candidates(words: list[SourceObservation], vocabulary: dict[str, tuple[str, ...]]) -> list[tuple[str, int, int, tuple[float, float, float, float], tuple[str, ...], str]]:
    candidates: list[tuple[str, int, int, tuple[float, float, float, float], tuple[str, ...], str]] = []
    for start in range(len(words)):
        for end in range(start + 1, len(words) + 1):
            text = " ".join(word.text or "" for word in words[start:end])
            value = match(text, vocabulary)
            boxes = [word.bbox for word in words[start:end] if word.bbox]
            if value is None or not boxes:
                continue
            bbox = (min(box[0] for box in boxes), min(box[1] for box in boxes), max(box[2] for box in boxes), max(box[3] for box in boxes))
            candidates.append((value, start, end, bbox, tuple(word.id for word in words[start:end]), text))
    # `match` deliberately accepts a known phrase inside surrounding text. For
    # geometry we need the smallest phrase at each occurrence, while retaining
    # repeated headers such as two distinct Abgänge columns.
    return [
        candidate for candidate in candidates
        if not any(
            other[0] == candidate[0]
            and other[1] >= candidate[1]
            and other[2] <= candidate[2]
            and (other[2] - other[1]) < (candidate[2] - candidate[1])
            for other in candidates
        )
    ]


def _pdf_group_for(center: float, groups: list[tuple[str, tuple[float, float, float, float], tuple[str, ...]]]) -> str | None:
    ordered = sorted(groups, key=lambda item: (item[1][0] + item[1][2]) / 2)
    for index, (group, bbox, _) in enumerate(ordered):
        group_center = (bbox[0] + bbox[2]) / 2
        left = float("-inf") if index == 0 else (((ordered[index - 1][1][0] + ordered[index - 1][1][2]) / 2) + group_center) / 2
        right = float("inf") if index + 1 == len(ordered) else (group_center + ((ordered[index + 1][1][0] + ordered[index + 1][1][2]) / 2)) / 2
        if left <= center < right:
            return group
    return None


def _pdf_field_for(group: str | None, direct: str | None, year: int | None, date_rank: int | None, date_count: int) -> tuple[str | None, str | None]:
    if group == "ahk":
        if direct == "zugang": return "zugang", None
        if direct == "abgang": return "abgang", None
        if direct == "umb": return "umb", None
        if year is not None and date_rank == 0: return "ahk_open", "opening"
        if year is not None and date_rank == date_count - 1: return "ahk_close", "closing"
    elif group == "cum_depreciation":
        if direct in {"total_afa_gj", "afa_plan"}: return "total_afa_gj", None
        if direct == "abgang": return "afa_abgang", None
        if direct == "zuschreibung": return "zuschreibung", None
        if direct == "umb": return "afa_umb", None
        if year is not None and date_rank == 0: return "kum_afa_open", "opening"
        if year is not None and date_rank == date_count - 1: return "kum_afa_close", "closing"
    elif group == "book_value" and year is not None:
        return ("bw_close", "closing") if date_rank == 0 else ("bw_prev", "prior_closing")
    if group is None and direct is None and year is not None and date_count >= 2:
        if date_rank == 0:
            return None, "opening"
        if date_rank == date_count - 1:
            return None, "closing"
    return direct, None


def _pdf_layout(amount: SourceObservation, observations: tuple[SourceObservation, ...]) -> _PdfLayout | None:
    if not amount.bbox:
        return None
    page = amount.locator.split(":", 2)[1]
    page_prefix = f"page:{page}:"
    rows = [observation for observation in observations if observation.kind == "logical_row" and observation.locator.startswith(page_prefix) and observation.bbox]
    visual_rows = [row for row in rows if row.metadata.get("structure_role") == "visual_row"]
    current = [row for row in visual_rows if amount.id in row.metadata.get("member_ids", ())]
    if len(current) != 1:
        return None
    # Azure table cells elect headers among cells of their own grid; page
    # words (title, unit, dates outside the table) stay context evidence.
    member_kind = "cell" if amount.kind == "cell" else "word"
    context_kinds = {"word", "cell"} if amount.kind == "cell" else {"word"}
    words = tuple(observation for observation in observations if observation.channel == amount.channel and observation.locator.startswith(page_prefix) and observation.kind in context_kinds and observation.text and observation.bbox)
    by_id = {word.id: word for word in words if word.kind == member_kind}
    header_options: list[tuple[tuple[int, int, float], SourceObservation, list[SourceObservation], list[tuple[str, int, int, tuple[float, float, float, float], tuple[str, ...], str]], list[tuple[str, int, int, tuple[float, float, float, float], tuple[str, ...], str]]]] = []
    for row in visual_rows:
        # Grid rows share their borders, so a cell row counts as above the
        # amount when its centre is; word rows keep the strict bottom test.
        row_bottom = None if row.bbox is None else ((row.bbox[1] + row.bbox[3]) / 2 if amount.kind == "cell" else row.bbox[3])
        if row_bottom is None or row_bottom >= amount.bbox[1]:
            continue
        row_words = sorted((by_id[item] for item in row.metadata.get("member_ids", ()) if item in by_id), key=lambda word: word.bbox[0])
        identities = _pdf_span_candidates(row_words, IDENTITIES)
        fields = _pdf_span_candidates(row_words, FIELDS)
        dates = sum(bool(re.search(r"\b(?:19|20)\d{2}\b", word.text or "")) for word in row_words)
        roles = {value for value, *_ in identities}
        if identities and (fields or dates):
            score = (int("asset_id" in roles) + int("source_name" in roles), len(roles), row.bbox[1])
            header_options.append((score, row, row_words, identities, fields))
    if not header_options:
        # The legacy layout extends each header span to the next one, which
        # would let an unlabelled grid column inherit its neighbour's field.
        return None if amount.kind == "cell" else _legacy_pdf_layout(amount, rows, words, current[0])
    _, header_row, header_words, identity_spans, field_spans = max(header_options, key=lambda item: item[0])
    assert header_row.bbox is not None

    group_rows: list[tuple[float, list[tuple[str, int, int, tuple[float, float, float, float], tuple[str, ...], str]]]] = []
    for row in visual_rows:
        if row.bbox is None or (row.bbox[1] + row.bbox[3]) / 2 >= (header_row.bbox[1] + header_row.bbox[3]) / 2:
            continue
        row_words = sorted((by_id[item] for item in row.metadata.get("member_ids", ()) if item in by_id), key=lambda word: word.bbox[0])
        spans = _pdf_span_candidates(row_words, HEADER_GROUPS)
        if spans:
            group_rows.append((row.bbox[1], spans))
    group_spans = max(group_rows, key=lambda item: (len(item[1]), item[0]))[1] if group_rows else []
    groups = [(value, bbox, evidence) for value, _, _, bbox, evidence, _ in group_spans]

    leaves: list[dict[str, object]] = []
    for direct, _, _, bbox, evidence, text in field_spans:
        if direct != "amount":
            leaves.append({"left": bbox[0], "bbox": bbox, "direct": direct, "year": _date_year(text), "text": text, "evidence": evidence})
    occupied = {(round(float(leaf["bbox"][0]), 3), round(float(leaf["bbox"][2]), 3)) for leaf in leaves}
    occupied_ids = {str(observation_id) for leaf in leaves for observation_id in leaf["evidence"]}
    for word in header_words:
        assert word.bbox is not None
        year = _date_year(word.text or "")
        key = (round(word.bbox[0], 3), round(word.bbox[2], 3))
        if year is not None and key not in occupied and word.id not in occupied_ids:
            leaves.append({"left": word.bbox[0], "bbox": word.bbox, "direct": None, "year": year, "text": word.text or "", "evidence": (word.id,)})
    leaves.sort(key=lambda item: float(item["left"]))
    for leaf in leaves:
        bbox = leaf["bbox"]
        leaf["group"] = _pdf_group_for((bbox[0] + bbox[2]) / 2, groups)
    ordered_groups = [value for value, _, _ in sorted(groups, key=lambda item: (item[1][0] + item[1][2]) / 2)]
    ordered_dates = [leaf for leaf in leaves if leaf["year"] is not None]
    # The supported asset-register hierarchy defines opening/closing date
    # leaves for each ordered metric group. When the complete pair structure is
    # visible, it is stronger evidence than the short centred parent caption,
    # whose glyph box does not span its edge columns.
    if ordered_groups and len(ordered_dates) == 2 * len(ordered_groups):
        for index, leaf in enumerate(ordered_dates):
            leaf["group"] = ordered_groups[index // 2]
    dates_by_group: dict[str | None, list[dict[str, object]]] = {}
    for leaf in leaves:
        if leaf["year"] is not None:
            dates_by_group.setdefault(leaf["group"], []).append(leaf)
    axes: list[_PdfAxis] = []
    for leaf in leaves:
        group_dates = dates_by_group.get(leaf["group"], [])
        date_rank = group_dates.index(leaf) if leaf in group_dates else None
        field, temporal = _pdf_field_for(leaf["group"], leaf["direct"], leaf["year"], date_rank, len(group_dates))
        group_evidence = next((evidence for value, _, evidence in groups if value == leaf["group"]), ())
        bbox = leaf["bbox"]
        axes.append(_PdfAxis(float(leaf["left"]), float(bbox[2]), field, str(leaf["year"]) if leaf["year"] else None, temporal, leaf["group"], str(leaf["text"]), tuple(dict.fromkeys((*leaf["evidence"], *group_evidence)))))
    if not axes:
        return None
    periods_by_field: dict[str, set[str]] = {}
    for axis in axes:
        if axis.field and axis.period:
            periods_by_field.setdefault(axis.field, set()).add(axis.period)
    numeric_data_right = max((
        word.bbox[2] for word in words
        if word.bbox and word.bbox[1] > header_row.bbox[3] and _looks_numeric(word.text or "")
    ), default=axes[-1].right)
    axes = [
        # A header cell is the exact column extent: a column without its own
        # field header must not inherit the field of the labelled column to its
        # left. Word headers keep extending to the next header.
        _PdfAxis(axis.left, axis.right if amount.kind == "cell" else (axes[index + 1].left if index + 1 < len(axes) else max(axis.right, numeric_data_right)), axis.field, axis.period, axis.temporal_role, axis.group, axis.header_text, axis.evidence_ids, bool(axis.field and len(periods_by_field.get(axis.field, ())) > 1))
        for index, axis in enumerate(axes)
    ]
    axes = [
        _PdfAxis(axis.left, axis.right, _pdf_compound_disposal_field(words, amount) if axis.field == "abgang" and _pdf_axis_has_reduction_counterevidence(axis, words, amount) else axis.field, axis.period, axis.temporal_role, axis.group, axis.header_text, axis.evidence_ids, axis.parallel_period)
        for axis in axes
    ]
    identity_axes = []
    ordered_identities = sorted(identity_spans, key=lambda item: item[3][0])
    for index, (role, _, _, bbox, evidence, _) in enumerate(ordered_identities):
        if identity_axes and role == identity_axes[-1].role and abs(bbox[0] - identity_axes[-1].left) <= 1:
            continue
        next_header_left = min((word.bbox[0] for word in header_words if word.bbox and word.bbox[0] > bbox[2]), default=axes[0].left)
        # A table cell already is the column extent; word headers need the gap midpoint.
        right = min(axes[0].left, bbox[2] if amount.kind == "cell" else (bbox[2] + next_header_left) / 2)
        identity_axes.append(_PdfIdentityAxis(float("-inf") if not identity_axes else bbox[0], right, role, evidence[0]))
    reporting_years = [int(year) for word in words if word.bbox and word.bbox[3] <= header_row.bbox[3] for year in _reported_periods(word.text or "")]
    reporting_period = str(max(reporting_years)) if reporting_years else None
    return _PdfLayout(current[0], words, tuple(visual_rows), tuple(axes), tuple(identity_axes), reporting_period, header_row.bbox[1], header_row.bbox[3])


def _legacy_pdf_layout(amount: SourceObservation, rows: list[SourceObservation], words: tuple[SourceObservation, ...], row: SourceObservation) -> _PdfLayout | None:
    assert amount.bbox is not None
    semantic_headers = [header for header in rows if header.metadata.get("structure_role") == "native_line" and header.bbox and header.bbox[3] < amount.bbox[1]]
    field_headers = sorted((header for header in semantic_headers if (header.metadata.get("semantic_field") or match(header.text or "", FIELDS))), key=lambda header: header.bbox[0])
    identity_headers = sorted((header for header in semantic_headers if (header.metadata.get("semantic_identity") or match(header.text or "", IDENTITIES))), key=lambda header: header.bbox[0])
    if not field_headers:
        return None
    axes = tuple(_PdfAxis(header.bbox[0], field_headers[index + 1].bbox[0] if index + 1 < len(field_headers) else header.bbox[2], str(header.metadata.get("semantic_field") or match(header.text or "", FIELDS)), str(_date_year(header.text or "")) if _date_year(header.text or "") else None, None, None, header.text or "", (header.id,)) for index, header in enumerate(field_headers))
    identity_axes = tuple(_PdfIdentityAxis(float("-inf") if index == 0 else header.bbox[0], identity_headers[index + 1].bbox[0] if index + 1 < len(identity_headers) else axes[0].left, str(header.metadata.get("semantic_identity") or match(header.text or "", IDENTITIES)), header.id) for index, header in enumerate(identity_headers))
    years = [int(year) for header in semantic_headers for year in _reported_periods(header.text or "")]
    return _PdfLayout(row, words, tuple(candidate for candidate in rows if candidate.metadata.get("structure_role") == "visual_row"), axes, identity_axes, str(max(years)) if years else None, min(header.bbox[1] for header in semantic_headers), max(header.bbox[3] for header in semantic_headers))


def _pdf_unit(amount: SourceObservation, layout: _PdfLayout, selected: _PdfAxis | None) -> tuple[str | None, tuple[str, ...]]:
    if not amount.bbox:
        return None, ()
    exact_units = [word for word in layout.words if word.bbox and re.search(r"(?:T?EUR|T€)", word.text or "", re.I) and word.bbox[3] < amount.bbox[1]]
    local: list[SourceObservation] = []
    if selected is not None:
        index = layout.axes.index(selected)
        right = selected.right
        local = [word for word in exact_units if word.bbox and word.bbox[1] >= layout.header_top and selected.left <= (word.bbox[0] + word.bbox[2]) / 2 <= right]
        header_values = ["TEUR" if re.search(r"\b(?:TEUR|T€)\b", selected.header_text, re.I) else "EUR"] if re.search(r"\b(?:T?EUR|T€)\b", selected.header_text, re.I) else []
    else:
        header_values = []
    local_values = header_values + [("TEUR" if re.search(r"\b(?:TEUR|T€)\b", word.text or "", re.I) else "EUR") for word in local]
    if local_values:
        values = set(local_values)
        return (next(iter(values)) if len(values) == 1 else None), tuple(word.id for word in local)
    global_units = [word for word in exact_units if word.bbox and word.bbox[3] < layout.header_top]
    values = {"TEUR" if re.search(r"\b(?:TEUR|T€)\b", word.text or "", re.I) else "EUR" for word in global_units}
    return (next(iter(values)) if len(values) == 1 else None), tuple(word.id for word in global_units)


def _compose_matrix_field(measure: str | None, axis: _PdfAxis, amount: SourceObservation) -> str | None:
    """Compose independent source row/column axes through a closed registry."""
    return compose_matrix_axes(measure, axis.field, axis.temporal_role)


def unit_from_text(text: str | None) -> str | None:
    """EUR or TEUR printed in the text; any other scale (Mio., Mrd.) or no currency abstains."""
    value = text or ""
    if re.search(r"\b(?:mio|million(?:en)?|mrd|milliarde(?:n)?)\b", value, re.I):
        return None
    if re.search(r"\b(?:tsd|tausend|teur|teuro)\b|T€", value, re.I):
        return "TEUR"
    return "EUR" if re.search(r"\bEUR\b|€", value, re.I) else None


def compose_matrix_axes(measure: str | None, field: str | None, temporal_role: str | None) -> str | None:
    """Compose a closed measure row by movement/period column intersection."""
    if measure is None:
        return field
    if measure == "ahk":
        if temporal_role == "opening":
            return "ahk_open"
        if temporal_role == "closing":
            return "ahk_close"
        return field if field in {"zugang", "abgang", "umb", "ahk_open", "ahk_close"} else None
    if measure == "cum_depreciation":
        if temporal_role == "opening":
            return "kum_afa_open"
        if temporal_role == "closing":
            return "kum_afa_close"
        return {
            "abgang": "afa_abgang",
            "umb": "afa_umb",
            "zuschreibung": "zuschreibung",
            "total_afa_gj": "total_afa_gj",
            "afa_plan": "afa_plan",
            "extraordinary": "extraordinary",
        }.get(field or "")
    if measure == "book_value":
        if temporal_role == "opening":
            return "bw_prev"
        if temporal_role == "closing":
            return "bw_close"
        return "abgang_bw" if field in {"abgang", "abgang_bw"} else None
    return None


def _pdf_axis_has_reduction_counterevidence(axis: _PdfAxis, words: tuple[SourceObservation, ...], amount: SourceObservation) -> bool:
    return any(
        word.bbox
        and word.bbox[3] < (amount.bbox[1] if amount.bbox else float("inf"))
        and axis.left <= (word.bbox[0] + word.bbox[2]) / 2 <= axis.right
        and "minderung" in norm(word.text or "")
        for word in words
    )


def _pdf_compound_disposal_field(words: tuple[SourceObservation, ...], amount: SourceObservation) -> str | None:
    source_basis = norm(" ".join(word.text or "" for word in words if word.bbox and word.bbox[3] < (amount.bbox[1] if amount.bbox else float("inf"))))
    if "netto anlage" in source_basis or "net asset" in source_basis:
        return "abgang_bw"
    if ("anschaff" in source_basis and "herstell" in source_basis) or "historical cost" in source_basis:
        return "abgang"
    return None


def _ocr_identity_text_has_noise(value: str | None) -> bool:
    return bool(value and any(token and not any(character.isalnum() for character in token) for token in value.split()))


def _pdf_ocr_identity_consistent(layout: _PdfLayout, current_asset_id: str) -> bool:
    """Require an OCR asset ID to agree with its source section's ID shape.

    A numeric heading without monetary members starts a section.  The data
    rows until the next such heading establish the repeated identifier width.
    A unique modal width is source counterevidence against an OCR token with a
    different width; no fixed confidence threshold or silent truncation is
    used.
    """
    if not current_asset_id.isdigit() or not layout.axes or not layout.identity_axes:
        return True
    asset_axes = [axis for axis in layout.identity_axes if axis.role == "asset_id"]
    if not asset_axes:
        return True
    asset_left = asset_axes[0].left
    asset_right = asset_axes[0].right
    page_rows: list[tuple[float, bool, str | None]] = []
    words_by_id = {word.id: word for word in layout.words}
    for row in sorted((candidate for candidate in layout.page_rows if candidate.bbox), key=lambda candidate: candidate.bbox[1]):
        members = [words_by_id[item] for item in row.metadata.get("member_ids", ()) if item in words_by_id and words_by_id[item].bbox]
        ids = [word.text or "" for word in members if asset_left <= (word.bbox[0] + word.bbox[2]) / 2 < asset_right and (word.text or "").isdigit()]
        has_amount = any(_looks_numeric(word.text or "") and (word.bbox[0] + word.bbox[2]) / 2 >= layout.axes[0].left for word in members)
        page_rows.append((row.bbox[1], has_amount, ids[0] if len(ids) == 1 else None))
    current_y = layout.row.bbox[1] if layout.row.bbox else 0
    headings = [index for index, (top, has_amount, value) in enumerate(page_rows) if top < current_y and not has_amount and value]
    start = headings[-1] if headings else -1
    end = next((index for index in range(start + 1, len(page_rows)) if page_rows[index][0] > current_y and not page_rows[index][1] and page_rows[index][2]), len(page_rows))
    widths = [len(value) for _, has_amount, value in page_rows[start + 1:end] if has_amount and value]
    if len(widths) < 2:
        return True
    counts = {width: widths.count(width) for width in set(widths)}
    best = max(counts.values())
    modes = [width for width, count in counts.items() if count == best]
    return len(modes) != 1 or len(current_asset_id) == modes[0]


def _looks_numeric(text: str) -> bool:
    return bool(re.fullmatch(r"\s*[+\-]?(?:\(?\s*)?[\d.,\u00a0\u202f ]+(?:\s*\)?|\s*[+\-])?(?:\s*(?:CR|DR))?\s*", text, re.I))


def _reported_periods(text: str) -> tuple[str, ...]:
    normalized = norm(text)
    years = tuple(re.findall(r"\b((?:19|20)\d{2})\b", text))
    if not years:
        return ()
    has_full_date = bool(re.search(r"\b\d{1,2}[./-]\d{1,2}[./-](?:19|20)\d{2}\b", text))
    has_report_label = any(label in normalized for label in ("jahresabschluss", "geschaftsjahr", "fiscal", "reporting period", "period ending", "stand zum", "per "))
    return years if has_full_date or has_report_label else ()


def _header_span(column: int, observation: SourceObservation) -> bool:
    merged = observation.metadata.get("merged_range")
    if isinstance(merged, str):
        min_col, _, max_col, _ = range_boundaries(merged)
        return min_col <= column <= max_col
    coordinate = observation.metadata.get("coordinate")
    if not isinstance(coordinate, str) and "!" in observation.locator:
        coordinate = observation.locator.rsplit("!", 1)[1]
    if not isinstance(coordinate, str):
        return False
    letters, _ = coordinate_from_string(coordinate)
    return column_index_from_string(letters) == column


def _spans_multiple_columns(observation: SourceObservation) -> bool:
    merged = observation.metadata.get("merged_range")
    if not isinstance(merged, str):
        return False
    min_col, _, max_col, _ = range_boundaries(merged)
    return max_col > min_col


def _header_group(text: str) -> str | None:
    return match(text, HEADER_GROUPS)


def _date_year(text: str) -> int | None:
    matches = re.findall(r"\b((?:19|20)\d{2})\b", text)
    return int(matches[-1]) if len(set(matches)) == 1 else None


def _xlsx_header_scope(amount_row: int, indexed: list[tuple[int, int, SourceObservation]]) -> tuple[int, int]:
    identity_headers = [(column, row) for column, row, observation in indexed if row < amount_row and observation.text and match(observation.text, IDENTITIES)]
    identity_rows = {row for _, row in identity_headers}
    if identity_rows:
        identity_header_row = max(identity_rows)
        previous_headers = sorted(row for row in identity_rows if row < identity_header_row)
        scope_start = min((row for _, row, _ in indexed), default=1)
        if previous_headers:
            previous_header = previous_headers[-1]
            intervening_data_rows = {
                row for _, row, observation in indexed
                if previous_header < row < identity_header_row
                and observation.text and _looks_numeric(observation.text)
            }
            if intervening_data_rows:
                scope_start = max(intervening_data_rows) + 1
        identity_columns = {column for column, row in identity_headers if row == identity_header_row}
        rows: dict[int, list[tuple[int, SourceObservation]]] = {}
        for column, row, observation in indexed:
            if row > identity_header_row:
                rows.setdefault(row, []).append((column, observation))
        data_rows = [
            row for row, values in rows.items()
            if any(column in identity_columns and observation.text for column, observation in values)
            and any(column not in identity_columns and _looks_numeric(observation.text or "") for column, observation in values)
        ]
        return scope_start, (min(data_rows) - 1 if data_rows else identity_header_row)
    semantic_rows = sorted({row for _, row, observation in indexed if row < amount_row and observation.text and (match(observation.text, FIELDS) or _header_group(observation.text))})
    cutoff = semantic_rows[-1] if semantic_rows else amount_row - 1
    previous_semantic = [row for row in semantic_rows if row < cutoff]
    scope_start = previous_semantic[-1] + 1 if previous_semantic and any(_looks_numeric(o.text or "") for _, r, o in indexed if previous_semantic[-1] < r < cutoff) else min((row for _, row, _ in indexed), default=1)
    return scope_start, cutoff


def _xlsx_header_cutoff(amount_row: int, indexed: list[tuple[int, int, SourceObservation]]) -> int:
    return _xlsx_header_scope(amount_row, indexed)[1]


def _xlsx_axis(column: int, amount_row: int, indexed: list[tuple[int, int, SourceObservation]], header_start: int, header_cutoff: int) -> tuple[str | None, str | None, str | None, tuple[str, ...]]:
    headers = [(c, r, o) for c, r, o in indexed if header_start <= r <= header_cutoff and o.text]
    covering = [(c, r, o) for c, r, o in headers if _header_span(column, o)]
    parents = [(r, o, _header_group(o.text or "")) for _, r, o in covering if _header_group(o.text or "")]
    parent = max(parents, key=lambda item: item[0]) if parents else None
    parent_row = parent[0] if parent else -1
    group = parent[2] if parent else None
    leaves = [(r, o) for _, r, o in covering if r > parent_row]
    leaf_row = max((r for r, _ in leaves), default=-1)
    leaf_observations = [o for r, o in leaves if r == leaf_row]
    leaf_text = " ".join(o.text or "" for o in leaf_observations)
    field = match(leaf_text, FIELDS) or match(" ".join(o.text or "" for _, _, o in covering), FIELDS)
    compound_text = norm(" ".join(o.text or "" for _, _, o in covering))
    if "abgang" in compound_text and "minderung" in compound_text:
        source_basis = norm(" ".join(o.text or "" for _, header_row, o in headers if header_row <= header_cutoff))
        if "netto anlage" in source_basis or "net asset" in source_basis:
            field = "abgang_bw"
        elif ("anschaff" in source_basis and "herstell" in source_basis) or "historical cost" in source_basis:
            field = "abgang"
        else:
            field = None
    temporal_role = None
    period_year = _date_year(leaf_text)
    if period_year is None:
        covering_years = {_date_year(o.text or "") for _, _, o in covering if _date_year(o.text or "") is not None}
        period_year = next(iter(covering_years)) if len(covering_years) == 1 else None
    if field is None and period_year is not None:
        dated_columns: dict[int, int] = {}
        for candidate_column in sorted({candidate for candidate, _, _ in headers}):
            years = {
                year for _, _, observation in headers
                if _header_span(candidate_column, observation)
                and (year := _date_year(observation.text or "")) is not None
            }
            if len(years) == 1:
                dated_columns[candidate_column] = next(iter(years))
        distinct_years = set(dated_columns.values())
        if len(distinct_years) > 1 and column in dated_columns:
            if dated_columns[column] == min(distinct_years):
                temporal_role = "opening"
            elif dated_columns[column] == max(distinct_years):
                temporal_role = "closing"
    evidence = [o.id for _, _, o in covering if o.text]
    if group:
        parent_observation = parent[1]
        merged = parent_observation.metadata.get("merged_range")
        if isinstance(merged, str):
            min_col, _, max_col, _ = range_boundaries(merged)
        else:
            min_col = max_col = column
        peer_dates: dict[int, int] = {}
        for peer_column in range(min_col, max_col + 1):
            peer_texts = [o.text or "" for _, r, o in headers if r > parent_row and _header_span(peer_column, o)]
            years = {_date_year(text) for text in peer_texts if _date_year(text) is not None}
            if len(years) == 1:
                peer_dates[peer_column] = next(iter(years))
        value = norm(leaf_text)
        if group == "ahk":
            if "zugang" in value or "zugänge" in value: field = "zugang"
            elif "abgang" in value or "abgänge" in value: field = "abgang"
            elif "umbuch" in value or "transfer" in value: field = "umb"
            elif column in peer_dates and column == min(peer_dates): field, temporal_role = "ahk_open", "opening"
            elif column in peer_dates and column == max(peer_dates): field, temporal_role = "ahk_close", "closing"
        elif group == "cum_depreciation":
            if "abgang" in value or "abgänge" in value: field = "afa_abgang"
            elif "zuschreib" in value or "write up" in value: field = "zuschreibung"
            elif "umbuch" in value: field = "afa_umb"
            elif "geschäftsjahr" in value or "geschaftsjahr" in value or "afa gj" in value: field = "total_afa_gj"
            # Additions to cumulative depreciation are the year's depreciation charge, never asset additions.
            elif "zugang" in value or "zugänge" in value: field = "total_afa_gj"
            elif column in peer_dates and column == min(peer_dates): field, temporal_role = "kum_afa_open", "opening"
            elif column in peer_dates and column == max(peer_dates): field, temporal_role = "kum_afa_close", "closing"
        elif group == "book_value" and column in peer_dates:
            newest = max(peer_dates.values())
            field = "bw_close" if peer_dates[column] == newest else "bw_prev"
            temporal_role = "closing" if field == "bw_close" else "prior_closing"
        if peer_dates:
            period_year = max(peer_dates.values())
    return field, str(period_year) if period_year else None, temporal_role, tuple(dict.fromkeys(evidence))


def _xlsx_record_role(row: int, indexed: list[tuple[int, int, SourceObservation]]) -> str:
    row_values = [(column, o) for column, current_row, o in indexed if current_row == row and o.text]
    values = [o for _, o in row_values]
    texts = [norm(o.text or "") for o in values]
    if any(text.startswith(("summe ", "subtotal ")) for text in texts):
        return "subtotal"
    if any(text.startswith(("gesamt ", "total ")) for text in texts):
        return "total"
    identity_columns = {column for column, header_row, header in indexed if header_row < row and header.text and match(header.text, IDENTITIES)}
    header_start, header_cutoff = _xlsx_header_scope(row, indexed)
    semantic_columns = {
        column for column, observation in row_values
        if observation.text
        and (field := _xlsx_axis(column, row, indexed, header_start, header_cutoff)[0]) not in {None, "amount"}
    }
    numeric_values = [o for column, o in row_values if column not in identity_columns and column in semantic_columns and _looks_numeric(o.text or "")]
    return "data" if numeric_values else "group_or_note"
