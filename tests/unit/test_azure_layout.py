"""Option C: Azure layout tables as primary evidence for scan/outline pages.

Every Azure response here is a REAL recorded prebuilt-layout response
(tests/fixtures/azure_responses) replayed through a loopback stub. The live
service is never called.
"""
from __future__ import annotations

import importlib.util
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread
import time
from urllib.parse import parse_qs, urlparse

import pymupdf as fitz
import pytest

from carthagent.acquire_pdf import acquire_pdf
from carthagent.numeric import parse_amount
from carthagent.ocr_azure import AzureDocumentIntelligenceOCR, AzureOCRConfig
from carthagent.policies import DEFAULT_BOUNDS
from carthagent.semantics import _pointer_for


ROOT = Path(__file__).resolve().parents[2]
RESPONSES = ROOT / "tests" / "fixtures" / "azure_responses"
KEY_ENV = "CARTHAGENT_TEST_AZURE_LAYOUT_KEY"
API_VERSION = "2024-11-30"


def recorded(name: str) -> dict:
    return json.loads((RESPONSES / f"{name}.json").read_text())


class LayoutStub(BaseHTTPRequestHandler):
    payload: dict | bytes = {}
    post_status = 202
    poll_status = 200
    running_polls = 0
    queries: list[dict[str, list[str]]] = []
    posts = 0
    gets = 0

    @classmethod
    def reset(cls, payload: dict | bytes) -> None:
        cls.payload = payload
        cls.post_status = 202
        cls.poll_status = 200
        cls.running_polls = 0
        cls.queries = []
        cls.posts = 0
        cls.gets = 0

    def do_POST(self) -> None:
        type(self).posts += 1
        type(self).queries.append(parse_qs(urlparse(self.path).query))
        self.rfile.read(int(self.headers["Content-Length"]))
        if type(self).post_status != 202:
            self.send_error(type(self).post_status)
            return
        self.send_response(202)
        self.send_header("Operation-Location", f"http://127.0.0.1:{self.server.server_port}/operations/1")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:
        type(self).gets += 1
        if type(self).poll_status != 200:
            self.send_error(type(self).poll_status)
            return
        payload = {"status": "running"} if type(self).gets <= type(self).running_polls or type(self).running_polls < 0 else type(self).payload
        body = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args: object) -> None:
        return


@pytest.fixture
def stub():
    LayoutStub.reset({})
    server = ThreadingHTTPServer(("127.0.0.1", 0), LayoutStub)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{server.server_port}"
    server.shutdown()
    thread.join()


@pytest.fixture(scope="module")
def documents(tmp_path_factory):
    directory = tmp_path_factory.mktemp("azure_sources")
    spec = importlib.util.spec_from_file_location("azure_cases", ROOT / "tests" / "fixtures" / "azure_cases.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.generate(directory)
    return directory


def provider(endpoint: str, **changes: object) -> AzureDocumentIntelligenceOCR:
    values: dict[str, object] = {
        "endpoint": endpoint, "api_version": API_VERSION, "key_env": KEY_ENV, "timeout_seconds": 5,
        "transient_retries": 0, "retry_delay_seconds": 0, "max_input_bytes": 50_000_000,
        "max_response_bytes": 5_000_000, "poll_interval_seconds": 0,
    }
    values.update(changes)
    return AzureDocumentIntelligenceOCR(AzureOCRConfig(**values))  # type: ignore[arg-type]


def acquire(path: Path, endpoint: str, **kwargs):
    return acquire_pdf(path, "doc", "hash", DEFAULT_BOUNDS, ocr_provider="azure_document_intelligence", azure_provider=provider(endpoint), **kwargs)


def cells_of(result):
    return [o for o in result.observations if o.kind == "cell" and o.channel == "pdf_ocr"]


def scan_doc(tmp_path: Path, pages: int) -> Path:
    """Blank raster pages with the corpus page size; geometry only, no text layer."""
    target = tmp_path / f"blank_{pages}.pdf"
    with fitz.open() as doc:
        for _ in range(pages):
            page = doc.new_page(width=842, height=595)
            pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 8, 8), False)
            pixmap.clear_with(255)
            page.insert_image(page.rect, stream=pixmap.tobytes("png"))
        doc.save(target)
    return target


def native_words(path: Path, text: str) -> list[tuple[float, float, float, float]]:
    with fitz.open(path) as doc:
        return [tuple(w[:4]) for w in doc[0].get_text("words") if w[4] == text]


def centre(box) -> tuple[float, float]:
    return (box[0] + box[2]) / 2, (box[1] + box[3]) / 2


def inside(point, box, tolerance: float = 0.0) -> bool:
    return box[0] - tolerance <= point[0] <= box[2] + tolerance and box[1] - tolerance <= point[1] <= box[3] + tolerance


def contains(outer, inner, tolerance: float = 2.0) -> bool:
    """The pointer box (a cell polygon) must enclose the printed lexeme's box."""
    return outer[0] - tolerance <= inner[0] and outer[1] - tolerance <= inner[1] and inner[2] <= outer[2] + tolerance and inner[3] <= outer[3] + tolerance


# ---------------------------------------------------------------- C1 cells


def test_each_azure_cell_becomes_a_cell_observation_with_grid_metadata(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_scanned.pdf", stub)
    cells = cells_of(result)
    assert len(cells) == 21
    required = {"row_index", "column_index", "row_span", "column_span", "cell_kind", "table_index", "page_index", "page_rotation", "page_width", "page_height", "confidence", "azure_cell_id", "verified"}
    for cell in cells:
        assert required <= set(cell.metadata), cell.metadata
        assert cell.locator.startswith("page:1:") and cell.bbox is not None
    assert {cell.metadata["cell_kind"] for cell in cells} >= {"columnHeader", "content"}
    rows = [o for o in result.observations if o.kind == "logical_row" and o.metadata.get("table_index") == 0 and o.metadata.get("structure_role") == "visual_row"]
    assert len(rows) == 3
    member_sets = [set(row.metadata["member_ids"]) for row in rows]
    assert set().union(*member_sets) == {cell.id for cell in cells}
    assert sum(len(m) for m in member_sets) == len(cells), "every cell belongs to exactly one visual row"
    regions = [o for o in result.observations if o.kind == "table_region" and o.metadata.get("table_index") == 0]
    assert len(regions) == 1 and list(regions[0].metadata["member_ids"]) == [row.id for row in sorted(rows, key=lambda r: r.metadata["row_index"])]
    # Azure words are still emitted: they are the cross-check source and stay pointer-capable.
    assert any(o.kind == "word" and o.channel == "pdf_ocr" and o.metadata.get("provider") == "azure_document_intelligence" for o in result.observations)


def test_sap_grid_spans_and_header_kinds_survive(stub, tmp_path, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("f05_sap_class_grid_03_p2"))
    result = acquire(scan_doc(tmp_path, 2), stub)
    cells = cells_of(result)
    assert len(cells) == 463
    assert all(cell.locator.startswith("page:2:") and cell.metadata["page_index"] == 1 for cell in cells)
    assert sum(cell.metadata["cell_kind"] == "columnHeader" for cell in cells) == 21
    spanning = [cell for cell in cells if cell.metadata["row_span"] > 1 or cell.metadata["column_span"] > 1]
    assert len(spanning) == 74
    rows = {o.id: o for o in result.observations if o.kind == "logical_row" and o.metadata.get("structure_role") == "visual_row" and o.metadata.get("table_index") == 0}
    assert len(rows) == 42
    owner = {}
    for row in rows.values():
        for member in row.metadata["member_ids"]:
            assert member not in owner, "a spanning cell must not be claimed by two rows"
            owner[member] = row
    for cell in cells:
        assert owner[cell.id].metadata["row_index"] == cell.metadata["row_index"]


@pytest.mark.parametrize(("name", "fixture", "shift"), [
    ("azure_scanned.pdf", "scanned_p1", (0, 0)),
    ("azure_scanned_rotated_90.pdf", "scanned_rotated_90_p1", (0, 0)),
    ("azure_scanned_cropped.pdf", "scanned_cropped_p1", (-15, -10)),
])
def test_numeric_cells_map_back_to_the_original_unrotated_page_points(stub, documents, monkeypatch, name, fixture, shift):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded(fixture))
    result = acquire(documents / name, stub)
    numeric = [cell for cell in cells_of(result) if cell.text and cell.text.strip() and cell.text.strip()[0].isdigit() and "," in cell.text]
    assert len(numeric) >= 8, [cell.text for cell in cells_of(result)]
    checked = 0
    for cell in numeric:
        targets = [(b[0] + shift[0], b[1] + shift[1], b[2] + shift[0], b[3] + shift[1]) for b in native_words(documents / "clear.pdf", cell.text.strip())]
        if not targets:
            continue
        # The cell polygon encloses exactly one printed occurrence of its text.
        assert sum(contains(cell.bbox, target) for target in targets) == 1, (name, cell.text, cell.bbox, targets)
        checked += 1
    assert checked >= 8


# ---------------------------------------------------------------- C2 cross-check


def test_cross_check_verifies_cells_whose_words_agree(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_scanned.pdf", stub)
    cells = cells_of(result)
    assert all(cell.metadata["verified"] is True for cell in cells), [(c.text, c.metadata.get("cross_check")) for c in cells if not c.metadata["verified"]]
    assert "page:1:azure_cells_verified:21" in result.diagnostics
    assert "page:1:azure_cells_unverified:0" in result.diagnostics
    # Each verified cell names the Azure words it was checked against.
    assert all(cell.metadata["member_word_ids"] for cell in cells if cell.text)


def test_cell_text_that_disagrees_with_its_words_is_unverified(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    cell = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["content"] == "1.000,00")
    cell["content"] = "7.000,00"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    bad = [c for c in cells_of(result) if c.text == "7.000,00"]
    assert len(bad) == 1 and bad[0].metadata["verified"] is False
    assert bad[0].metadata["cross_check"] == "text_mismatch"
    assert bad[0].text == "7.000,00", "unverified cells keep their text"
    assert "page:1:azure_cells_unverified:1" in result.diagnostics


def test_numeric_looking_cell_that_does_not_parse_is_unverified(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    result_raw = payload["analyzeResult"]
    cell = next(c for c in result_raw["tables"][0]["cells"] if c["content"] == "1.000,00")
    offset = cell["spans"][0]["offset"]
    word = next(w for w in result_raw["pages"][0]["words"] if w["span"]["offset"] == offset)
    cell["content"] = word["content"] = "1.0.0,0,0"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    bad = [c for c in cells_of(result) if c.text == "1.0.0,0,0"]
    assert len(bad) == 1 and bad[0].metadata["verified"] is False
    assert bad[0].metadata["cross_check"] == "numeric_unparseable"


def test_word_claimed_by_two_overlapping_cells_verifies_neither(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    cells = payload["analyzeResult"]["tables"][0]["cells"]
    source = next(c for c in cells if c["content"] == "1.000,00")
    neighbour = next(c for c in cells if c["rowIndex"] == source["rowIndex"] and c["columnIndex"] == source["columnIndex"] + 1)
    neighbour["boundingRegions"][0]["polygon"] = list(source["boundingRegions"][0]["polygon"])
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    by_position = {(c.metadata["row_index"], c.metadata["column_index"]): c for c in cells_of(result)}
    first = by_position[(source["rowIndex"], source["columnIndex"])]
    second = by_position[(neighbour["rowIndex"], neighbour["columnIndex"])]
    assert first.metadata["verified"] is False and second.metadata["verified"] is False
    assert first.metadata["cross_check"] == "word_in_multiple_cells"


def test_real_sap_grid_cross_check_is_consistent(stub, tmp_path, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("f05_sap_class_grid_03_p2"))
    result = acquire(scan_doc(tmp_path, 2), stub)
    cells = cells_of(result)
    words = {o.id: o for o in result.observations if o.kind == "word" and o.channel == "pdf_ocr"}
    for cell in cells:
        joined = " ".join(" ".join(words[w].text.split()) for w in cell.metadata["member_word_ids"])
        same_text = " ".join((cell.text or "").split()) == " ".join(joined.split())
        if not same_text:
            assert cell.metadata["verified"] is False, cell
        if cell.metadata["verified"]:
            assert same_text
            try:
                parse_amount(cell.text or "")
            except ValueError:
                continue
            assert parse_amount(cell.text).source_signed_value == parse_amount(joined).source_signed_value
    verified = sum(cell.metadata["verified"] for cell in cells)
    assert f"page:2:azure_cells_verified:{verified}" in result.diagnostics
    assert f"page:2:azure_cells_unverified:{len(cells) - verified}" in result.diagnostics


# ---------------------------------------------------------------- C3 provider policy


def test_native_pages_are_never_sent_to_azure(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "clear.pdf", stub)
    assert LayoutStub.posts == 0
    assert not cells_of(result)
    assert not any(o.metadata.get("provider") == "azure_document_intelligence" for o in result.observations)


def test_only_scan_and_outline_pages_are_requested(stub, documents, tmp_path, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    target = tmp_path / "native_then_scan.pdf"
    with fitz.open(documents / "clear.pdf") as doc, fitz.open(scan_doc(tmp_path, 2)) as scans:
        doc.insert_pdf(scans, from_page=1, to_page=1)
        doc.save(target)
    LayoutStub.reset(recorded("f05_sap_class_grid_03_p2"))
    result = acquire(target, stub)
    assert LayoutStub.posts == 1
    assert LayoutStub.queries[0]["pages"] == ["2"]
    assert cells_of(result) and all(c.metadata["page_index"] == 1 for c in cells_of(result))
    assert any(o.channel == "pdf_native" for o in result.observations)


def test_outline_page_routes_to_azure(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_outline_only.pdf", stub)
    assert LayoutStub.posts == 1 and LayoutStub.queries[0]["pages"] == ["1"]
    assert len(cells_of(result)) == 21


def test_mistral_is_not_invoked_under_the_azure_provider(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")

    class Forbidden:
        def analyze(self, *args, **kwargs):  # pragma: no cover - must not run
            raise AssertionError("mistral must not be called")

    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_scanned.pdf", stub, mistral_provider=Forbidden())
    assert cells_of(result)


@pytest.mark.parametrize("failure", ["post_429", "post_503", "poll_500", "timeout"])
def test_azure_failure_falls_back_to_tesseract(stub, documents, monkeypatch, failure):
    import shutil
    if not shutil.which("tesseract"):
        pytest.skip("Tesseract is required for the fallback path")
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    if failure == "post_429":
        LayoutStub.post_status = 429
    elif failure == "post_503":
        LayoutStub.post_status = 503
    elif failure == "poll_500":
        LayoutStub.poll_status = 500
    else:
        LayoutStub.running_polls = -1
    azure = provider(stub, timeout_seconds=0.5 if failure == "timeout" else 5)
    result = acquire_pdf(documents / "azure_scanned.pdf", "doc", "hash", DEFAULT_BOUNDS, ocr_provider="azure_document_intelligence", azure_provider=azure)
    assert LayoutStub.posts == 1
    assert any(d.startswith("azure_unavailable_fallback_tesseract:") for d in result.diagnostics), result.diagnostics
    assert not cells_of(result)
    assert any(o.channel == "pdf_ocr" and ":ocr:" in o.id for o in result.observations), "Tesseract words must be present"


def test_azure_wait_is_bounded_by_the_post_mapper_reserve(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    # Less time left than the reserve: Azure must not even be started.
    deadline = time.monotonic() + DEFAULT_BOUNDS.post_mapper_reserve_seconds - 1
    result = acquire(documents / "azure_scanned.pdf", stub, deadline=deadline, ocr_enabled=True)
    assert LayoutStub.posts == 0
    assert any(d.startswith("azure_unavailable_fallback_tesseract:") and "budget" in d for d in result.diagnostics), result.diagnostics


def test_client_sends_requested_pages_as_ranges(stub, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    provider(stub).analyze(b"%PDF-x", pages=(1, 2, 3, 5, 7, 8))
    assert LayoutStub.queries[0]["pages"] == ["1-3,5,7-8"]
    with pytest.raises(ValueError):
        provider(stub).analyze(b"%PDF-x", pages=(0,))


# ---------------------------------------------------------------- C4 pointers


def test_repeated_equal_amounts_on_a_scan_point_at_their_own_occurrence(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("repeated_values_scanned_p1"))
    result = acquire(documents / "azure_repeated_values_scanned.pdf", stub)
    cells = cells_of(result)
    with fitz.open(documents / "repeated_values.pdf") as doc:
        words = doc[0].get_text("words")
        page = doc[0]
        expected = {}
        for identity in ("000101", "000202"):
            label = next(w for w in words if w[4] == identity)
            y = (label[1] + label[3]) / 2
            row = sorted((w for w in words if w[4] == "1.000,00" and abs((w[1] + w[3]) / 2 - y) < 2), key=lambda w: w[0])
            assert len(row) == 5
            expected[identity] = [tuple(w[:4]) for w in row]
        occurrences = [box for boxes in expected.values() for box in boxes]
        by_row: dict[int, list] = {}
        for cell in cells:
            by_row.setdefault(cell.metadata["row_index"], []).append(cell)
        checked = 0
        for row_cells in by_row.values():
            identity = next((c.text.strip() for c in row_cells if c.metadata["column_index"] == 0 and c.text), None)
            if identity not in expected:
                continue
            amounts = sorted((c for c in row_cells if c.text == "1.000,00"), key=lambda c: c.metadata["column_index"])
            assert len(amounts) == 5
            for cell, target in zip(amounts, expected[identity]):
                pointer = _pointer_for(cell)
                assert pointer["coordinate_space"] == "pymupdf_unrotated_points_top_left"
                # Deliberate (part B): a verified cell points at its printed lexeme;
                # the cell polygon stays in the pointer for a later convention switch.
                assert cell.metadata["verified"] is True
                assert pointer["bbox"] == list(cell.metadata["lexeme_bbox"])
                assert pointer["cell_polygon"] == list(cell.metadata["cell_polygon"]) == list(cell.bbox)
                # The lexeme box is tighter than the PDF glyph box: compare by centre.
                hits = [box for box in occurrences if inside(centre(pointer["bbox"]), box, 1)]
                assert hits == [target], (identity, cell.metadata["column_index"], pointer["bbox"], hits)
                # A second read of the original page at the pointer box finds the lexeme.
                assert "1.000,00" in page.get_textbox(fitz.Rect(pointer["bbox"])).replace(" ", "")
                checked += 1
        assert checked == 10


def test_page_azure_did_not_return_falls_back_with_a_diagnostic(stub, tmp_path, monkeypatch):
    # The free tier analyses only the first two pages; a requested page that
    # comes back missing must be visible, not silently handed to Tesseract.
    import shutil
    if not shutil.which("tesseract"):
        pytest.skip("Tesseract is required for the fallback path")
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("f05_sap_class_grid_03_p2"))
    result = acquire(scan_doc(tmp_path, 2), stub)
    assert LayoutStub.queries[0]["pages"] == ["1-2"]
    assert "page:1:azure_page_missing_fallback_tesseract" in result.diagnostics
    assert not any(d.startswith("page:2:azure_page_missing") for d in result.diagnostics)


def test_verified_cells_carry_the_printed_lexeme_box(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    payload = recorded("scanned_p1")
    bad = next(c for c in payload["analyzeResult"]["tables"][0]["cells"] if c["content"] == "2.000,00")
    bad["content"] = "9.000,00"
    LayoutStub.reset(payload)
    result = acquire(documents / "azure_scanned.pdf", stub)
    cells = cells_of(result)
    for cell in cells:
        if not cell.metadata["verified"] or not cell.text:
            assert cell.metadata.get("lexeme_bbox") is None
            continue
        lexeme = cell.metadata["lexeme_bbox"]
        assert contains(cell.bbox, lexeme, 0.01)
        targets = native_words(documents / "clear.pdf", cell.text.split()[0])
        if cell.text[0].isdigit() and targets:
            assert sum(inside(centre(t), lexeme, 2) for t in targets) == 1, (cell.text, lexeme, targets)
    assert any(c.text == "9.000,00" and c.metadata.get("lexeme_bbox") is None for c in cells)


# ---------------------------------------------------------------- deskew and recorded boxes (round 3, part A)


def _skewed_page(angle, points, *, width: float = 11.6944, height: float = 8.2639):
    from carthagent.ocr_azure import OCRPage, OCRTextSpan, OCRWord
    polygon = tuple(value for point in points for value in point)
    return OCRPage(1, width, height, "inch", angle, (OCRWord("1.000,00", polygon, 0.99, OCRTextSpan(0, 8)),))


def _rotate_clockwise(x: float, y: float, degrees: float, cx: float, cy: float) -> tuple[float, float]:
    import math
    t = math.radians(degrees)
    dx, dy = x - cx, y - cy
    return cx + dx * math.cos(t) - dy * math.sin(t), cy + dx * math.sin(t) + dy * math.cos(t)


@pytest.mark.parametrize(("angle", "expected_deskew"), [(1.2, 1.2), (-0.9, -0.9), (None, 0.0)])
def test_azure_skew_is_undone_about_the_page_centre(angle, expected_deskew):
    from carthagent.acquire_pdf import _azure_page_observations
    width, height = 11.6944, 8.2639
    # A word far right of the centre: known upright position (9.3 in, 2.075 in).
    upright = [(9.0, 2.0), (9.6, 2.0), (9.6, 2.15), (9.0, 2.15)]
    printed = [_rotate_clockwise(x, y, expected_deskew, width / 2, height / 2) for x, y in upright]
    with fitz.open() as doc:
        page = doc.new_page(width=width * 72, height=height * 72)
        observations, _ = _azure_page_observations(_skewed_page(angle, printed), page, "doc", "hash", 0)
    word = next(o for o in observations if o.kind == "word")
    assert word.metadata["deskew_degrees"] == pytest.approx(expected_deskew)
    x0, y0, x1, y1 = word.bbox
    assert ((x0 + x1) / 2, (y0 + y1) / 2) == pytest.approx((9.3 * 72, 2.075 * 72), abs=0.05)
    if expected_deskew:
        # Uncorrected, the printed position at this x is several points off.
        assert abs(sum(y for _, y in printed) / 4 * 72 - 2.075 * 72) > 1.5


def test_quarter_turn_page_keeps_only_the_residual_skew(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_rotated_90_p1"))
    result = acquire(documents / "azure_scanned_rotated_90.pdf", stub)
    angles = {o.metadata["deskew_degrees"] for o in result.observations if o.metadata.get("provider") == "azure_document_intelligence" and "deskew_degrees" in o.metadata}
    # Azure reports 89.9: the /Rotate quarter turn (handled by PageTransform) plus -0.1 skew.
    assert len(angles) == 1 and angles.pop() == pytest.approx(89.90962219238281 - 90)


def test_cells_record_both_boxes_and_the_deskew(stub, documents, monkeypatch):
    monkeypatch.setenv(KEY_ENV, "k")
    LayoutStub.reset(recorded("scanned_p1"))
    result = acquire(documents / "azure_scanned.pdf", stub)
    angle = recorded("scanned_p1")["analyzeResult"]["pages"][0]["angle"]
    for cell in cells_of(result):
        assert list(cell.metadata["cell_polygon"]) == list(cell.bbox), "the cell polygon box stays recorded"
        assert cell.metadata["deskew_degrees"] == pytest.approx(angle)
        if cell.metadata["verified"] and cell.text:
            lexeme = cell.metadata["lexeme_bbox"]
            assert contains(cell.bbox, lexeme, 0.01)
