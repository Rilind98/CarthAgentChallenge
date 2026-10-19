"""Pipeline-level Option C checks with a recorded Azure response replayed on loopback."""
from __future__ import annotations

from collections import Counter
from decimal import Decimal
from http.server import ThreadingHTTPServer
import importlib.util
import json
from pathlib import Path
import sys
from threading import Thread

import pymupdf
import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tests" / "unit"))
from test_azure_layout import KEY_ENV, LayoutStub, recorded  # noqa: E402
from test_pipeline import evidence_for  # noqa: E402

ALLOWED = {"000101": {"ahk_open": 1000, "zugang": 120, "abgang": 20, "ahk_close": 1100, "bw_close": 800},
           "000202": {"ahk_open": 2000, "zugang": 340, "abgang": 40, "ahk_close": 2300, "bw_close": 1600}}


@pytest.fixture(scope="module")
def documents(tmp_path_factory):
    directory = tmp_path_factory.mktemp("azure_pipeline_sources")
    spec = importlib.util.spec_from_file_location("azure_cases", ROOT / "tests" / "fixtures" / "azure_cases.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.generate(directory)
    return directory


@pytest.fixture
def azure_config(tmp_path, monkeypatch):
    LayoutStub.reset({})
    server = ThreadingHTTPServer(("127.0.0.1", 0), LayoutStub)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    monkeypatch.setenv(KEY_ENV, "loopback-only")
    config = tmp_path / "azure.toml"
    config.write_text(
        "[ocr]\nenabled = true\nprovider = \"azure_document_intelligence\"\n\n[ocr.azure]\n"
        f"endpoint = \"http://127.0.0.1:{server.server_port}\"\napi_version = \"2024-11-30\"\nkey_env = \"{KEY_ENV}\"\n"
        "timeout_seconds = 10\ntransient_retries = 0\nretry_delay_seconds = 0\npoll_interval_seconds = 0\n"
        "max_input_bytes = 50000000\nmax_response_bytes = 5000000\n"
    )
    yield config
    server.shutdown()
    thread.join()


def _display(value: Decimal) -> str:
    return f"{value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


@pytest.mark.parametrize(("name", "fixture", "shift"), [
    ("azure_scanned.pdf", "scanned_p1", (0, 0)),
    ("azure_scanned_rotated_90.pdf", "scanned_rotated_90_p1", (0, 0)),
    ("azure_scanned_cropped.pdf", "scanned_cropped_p1", (-15, -10)),
])
def test_azure_scan_facts_are_correct_unique_and_point_at_the_printed_value(documents, tmp_path, azure_config, name, fixture, shift):
    LayoutStub.reset(recorded(fixture))
    _, data = evidence_for(documents / name, tmp_path / "out", config=azure_config)
    assert LayoutStub.posts == 1
    assert not any(d.startswith("azure_unavailable_fallback_tesseract") for d in data["diagnostics"]), data["diagnostics"]
    assert any(":azure_cells_verified:" in d for d in data["diagnostics"]), data["diagnostics"]
    facts = data["facts"]
    # Two assets x five fields; a printed value must not appear twice (word and cell).
    assert len(facts) == 10, [(f["asset_identity"]["asset_or_class"], f["field"]) for f in facts]
    keys = Counter((f["asset_identity"]["asset_or_class"], f["field"]) for f in facts)
    assert all(count == 1 for count in keys.values()), keys
    with pymupdf.open(documents / "clear.pdf") as original:
        words = original[0].get_text("words")
    for fact in facts:
        identity = fact["asset_identity"]["asset_or_class"]
        value = Decimal(fact["amount"]["normalized_base"])
        assert identity in ALLOWED and fact["field"] in ALLOWED[identity], fact
        assert value == ALLOWED[identity][fact["field"]], fact
        pointer = fact["pointer"]
        assert pointer["kind"] == "pdf" and pointer["coordinate_space"] == "pymupdf_unrotated_points_top_left"
        box = pointer["bbox"]
        # The pointer (an Azure word or an Azure cell polygon) must enclose the
        # centre of exactly one printed occurrence of this value on the page.
        targets = [(w[0] + shift[0], w[1] + shift[1], w[2] + shift[0], w[3] + shift[1]) for w in words if w[4] == _display(value)]
        enclosed = [t for t in targets if box[0] - 2 <= (t[0] + t[2]) / 2 <= box[2] + 2 and box[1] - 2 <= (t[1] + t[3]) / 2 <= box[3] + 2]
        assert len(enclosed) == 1, (name, fact["field"], box, targets)


def test_native_document_is_identical_under_the_azure_provider(documents, tmp_path, azure_config):
    LayoutStub.reset(recorded("scanned_p1"))
    _, with_azure = evidence_for(documents / "clear.pdf", tmp_path / "azure", config=azure_config)
    _, default = evidence_for(documents / "clear.pdf", tmp_path / "default")
    assert LayoutStub.posts == 0

    def summary(data):
        return sorted((f["asset_identity"]["asset_or_class"], f["field"], f["amount"]["normalized_base"], json.dumps(f["pointer"]["bbox"])) for f in data["facts"])

    assert summary(with_azure) == summary(default) and summary(default)
