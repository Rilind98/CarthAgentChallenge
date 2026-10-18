from __future__ import annotations

from dataclasses import replace
from decimal import Decimal
from pathlib import Path

from hypothesis import given, strategies as st

from carthagent.geometry import PageTransform
from carthagent.models import SourceObservation
from carthagent.numeric import parse_amount
from carthagent.pipeline import _document_id, _observation_chunks
from carthagent.policies import DEFAULT_BOUNDS
from carthagent.store import Ledger


@given(st.integers(min_value=-10**15, max_value=10**15), st.integers(min_value=0, max_value=6))
def test_numeric_plain_decimal_replay(value: int, scale_power: int):
    scale = Decimal(10) ** scale_power
    parsed = parse_amount(str(value), scale=scale)
    assert parsed.source_signed_value == Decimal(value)
    assert parsed.normalized_value == Decimal(value) * scale


@given(
    st.floats(min_value=1, max_value=2000, allow_nan=False, allow_infinity=False),
    st.floats(min_value=1, max_value=2000, allow_nan=False, allow_infinity=False),
    st.sampled_from([0, 90, 180, 270]),
    st.floats(min_value=0.25, max_value=8, allow_nan=False, allow_infinity=False),
    st.data(),
)
def test_coordinate_transform_roundtrip(width: float, height: float, rotation: int, scale: float, data):
    x0 = data.draw(st.floats(min_value=0, max_value=width, allow_nan=False, allow_infinity=False))
    x1 = data.draw(st.floats(min_value=x0, max_value=width, allow_nan=False, allow_infinity=False))
    y0 = data.draw(st.floats(min_value=0, max_value=height, allow_nan=False, allow_infinity=False))
    y1 = data.draw(st.floats(min_value=y0, max_value=height, allow_nan=False, allow_infinity=False))
    transform = PageTransform(rotation=rotation, scale=scale, page_width=width, page_height=height)
    assert transform.roundtrip_ok((x0, y0, x1, y1), tolerance=1e-7)


@given(st.integers(min_value=1, max_value=100), st.integers(min_value=2, max_value=20), st.data())
def test_pdf_chunk_partition_covers_every_page(count: int, chunk_size: int, data):
    overlap = data.draw(st.integers(min_value=0, max_value=chunk_size - 1))
    observations = tuple(SourceObservation(f"d:p:{page}", "d", "a" * 64, "word", f"page:{page + 1}:word:0", "pdf_native", "x", metadata={"page_index": page}) for page in range(count))
    bounds = replace(DEFAULT_BOUNDS, pdf_chunk_pages=chunk_size, page_overlap=overlap)
    chunks = _observation_chunks(observations, bounds)
    assert {item.id for chunk in chunks for item in chunk} == {item.id for item in observations}
    assert all(len({item.metadata["page_index"] for item in chunk}) <= chunk_size for chunk in chunks)


@given(st.text(alphabet=st.characters(whitelist_categories=("Ll", "Lu", "Nd")), min_size=1, max_size=30), st.binary(min_size=32, max_size=32))
def test_document_and_cache_identity_are_stable_and_versioned(stem: str, digest: bytes):
    source_hash = digest.hex()
    path = Path(stem + ".xlsx")
    assert _document_id(path, source_hash) == _document_id(path, source_hash)
    first = Ledger.cache_key(source_hash, "region", {"mode": "x"}, {"parser": "1"})
    assert first == Ledger.cache_key(source_hash, "region", {"mode": "x"}, {"parser": "1"})
    assert first != Ledger.cache_key(source_hash, "region", {"mode": "x"}, {"parser": "2"})


def test_dependency_invalidation_reaches_binding_snapshot_and_export(tmp_path):
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    with ledger.transaction():
        for record_id, kind in (("obs", "observation"), ("fact", "fact"), ("snapshot", "snapshot"), ("binding", "output_binding"), ("export", "export")):
            ledger.append(record_id, "d", kind, {"id": record_id})
        ledger.depend("fact", "obs", "semantic")
        ledger.depend("snapshot", "fact", "semantic_snapshot")
        ledger.depend("binding", "fact", "output_projection")
        ledger.depend("binding", "snapshot", "snapshot_projection")
        ledger.depend("export", "binding", "publication")
    invalidated = ledger.invalidate("obs", "semantic")
    assert invalidated == {"fact", "snapshot", "binding", "export"}
    assert ledger.is_invalidated("export")
    ledger.close()
