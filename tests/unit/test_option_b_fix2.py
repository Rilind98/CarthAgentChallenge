"""Fix round 2 cost/coverage items: compact request payload (R2-3), prompt rules (R2-2), split-request coverage (R2-7)."""
from __future__ import annotations

import importlib.util
from pathlib import Path

from carthagent.acquire_xlsx import acquire_xlsx
from carthagent.mapper import MAPPER_PROMPT_VERSION, MapperConfig, OpenAICompatibleMapper
from carthagent.policies import DEFAULT_BOUNDS

ROOT = Path(__file__).resolve().parents[2]
_spec = importlib.util.spec_from_file_location("fix2_fixtures", ROOT / "tests" / "fixtures" / "build_option_b.py")
fixtures = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(fixtures)
ALWAYS = {"id", "kind", "text", "numeric_value", "row_block_id", "region_id", "locator_kind"}
GEOMETRY = {"sheet", "row", "column", "normalized_bbox"}


def _payload(observations):
    mapper = OpenAICompatibleMapper(MapperConfig(True, "http://127.0.0.1:1/v1/responses", "m", 1, 1))
    data, body = mapper._request_body("d", observations)
    return data, body, mapper


def test_r2_3_payload_drops_empty_keys_but_keeps_ids_and_geometry(tmp_path):
    observations = acquire_xlsx(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), "d", "a" * 64, DEFAULT_BOUNDS).observations
    data, body, _ = _payload(observations)
    for item in data:
        assert ALWAYS <= set(item), item
        assert not [key for key, value in item.items() if key not in ALWAYS | GEOMETRY and (value is None or value == [])], item
    cell = next(item for item in data if item["id"] == "d:xlsx:Anlagengitter!B20")
    assert (cell["sheet"], cell["row"], cell["column"], cell["numeric_value"], cell["text"]) == ("Anlagengitter", 20, 2, True, None)
    block = next(item for item in data if item["kind"] == "record_block")
    assert block["member_ids"] and block["identity_row_id"]


def test_r2_2_prompt_carries_the_structure_rules_and_a_new_version():
    _, body, mapper = _payload(())
    system = body["input"][0]["content"][0]["text"]
    for rule in ("empty header cell", "source_name", "record date", "total_afa_gj", "afa_abgang", "identity row"):
        assert rule in system, rule
    assert MAPPER_PROMPT_VERSION == "compact-region-plan-v12" and mapper.fingerprint()["prompt_version"] == MAPPER_PROMPT_VERSION


def test_r2_7_identical_claims_from_overlapping_requests_are_one_hypothesis():
    from carthagent.models import AssociationClaim, AssociationKind, HypothesisSet
    from carthagent.pipeline import _deduplicate_mapped
    claim = AssociationClaim(AssociationKind.OTHER, ("r",), ("r",), (), False, "ahk_open", record_identity_row_id="i")
    other = AssociationClaim(AssociationKind.OTHER, ("r",), ("r",), (), False, "kum_afa_open", record_identity_row_id="i")
    unknown = AssociationClaim(AssociationKind.UNKNOWN_UNMODELED, ())
    same = (HypothesisSet("d:mapper:a", "d", "x", (claim, unknown)), HypothesisSet("d:mapper:b", "d", "x", (claim, unknown)))
    assert len(_deduplicate_mapped(same)) == 1
    assert _deduplicate_mapped((*same, HypothesisSet("d:mapper:c", "d", "x", (other, unknown)))) == ()


def test_r2_7_every_row_window_carries_the_header_band_above_the_first_record(tmp_path):
    from dataclasses import replace
    from carthagent.pipeline import _observation_chunks
    observations = acquire_xlsx(fixtures.sap_block_unlabelled(tmp_path / "sap.xlsx"), "d", "a" * 64, DEFAULT_BOUNDS).observations
    chunks = _observation_chunks(observations, replace(DEFAULT_BOUNDS, row_chunk_size=8, row_overlap=4))
    header_band = {f"d:xlsx:Anlagengitter:row:{n}" for n in (15, 16, 17, 18)}
    assert len(chunks) > 2 and all(header_band <= {o.id for o in chunk} for chunk in chunks)
