from __future__ import annotations

from hashlib import sha256
from pathlib import Path

from openpyxl import Workbook
import pytest

from carthagent.acquire_xlsx import acquire_xlsx
from carthagent.policies import DEFAULT_BOUNDS
from carthagent.structure import FIELDS, ContextResolver, match


@pytest.mark.parametrize(("header", "expected"), [
    # A period qualifier next to a movement word names the movement, never current-year depreciation.
    ("Zugänge im Geschäftsjahr", "zugang"),
    ("Abgänge Geschäftsjahr", "abgang"),
    ("Umbuchungen im Geschäftsjahr", "umb"),
    ("Zuschreibungen des Geschäftsjahres", "zuschreibung"),
    # Depreciation wording keeps its meaning.
    ("Abschreibungen des Geschäftsjahres", "total_afa_gj"),
    ("AfA Geschäftsjahr", "total_afa_gj"),
    ("Geschäftsjahr", "total_afa_gj"),
])
def test_movement_words_outrank_period_qualifier(header, expected):
    assert match(header, FIELDS) == expected


def test_xlsx_group_branch_keeps_movement_under_depreciation_group(tmp_path: Path):
    source = tmp_path / "spiegel.xlsx"
    workbook = Workbook(); sheet = workbook.active; sheet.title = "Spiegel"
    sheet["A1"] = "Anlagenspiegel Wirtschaftsjahr 2025 EUR"
    sheet["B2"] = "Kumulierte Abschreibungen"; sheet.merge_cells("B2:C2")
    sheet.append(["Anl.-Nr.", "Zuschreibungen im Geschäftsjahr", "Geschäftsjahr"])
    sheet.append(["27001", 5, 7])
    workbook.save(source)
    acquired = acquire_xlsx(source, "d", sha256(source.read_bytes()).hexdigest(), DEFAULT_BOUNDS)
    observations = acquired.observations
    resolver = ContextResolver(observations)
    by_locator = {o.locator: o for o in observations}
    writeup = resolver.resolve(by_locator["Spiegel!B4"], observations)
    depreciation = resolver.resolve(by_locator["Spiegel!C4"], observations)
    assert writeup.business is not None and writeup.business.field == "zuschreibung"
    assert depreciation.business is not None and depreciation.business.field == "total_afa_gj"


@pytest.mark.parametrize(("header", "expected"), [
    # R3-5: an opening/closing suffix names the cumulative balance, never the year's charge.
    ("AfA GJ-Beg", "kum_afa_open"), ("AfA GJ Beginn", "kum_afa_open"), ("AfA Anfang", "kum_afa_open"),
    ("AfA GJ-Ende", "kum_afa_close"), ("AfA GJ End", "kum_afa_close"),
    ("AfA GJ", "total_afa_gj"), ("AfA des Jahres", None),
])
def test_depreciation_suffix_names_the_balance_not_the_charge(header, expected):
    assert match(header, FIELDS) == expected


@pytest.mark.parametrize(("header", "expected"), [
    ("UNr", "asset_sub_id"), ("U-Nr.", "asset_sub_id"), ("Sub", "asset_sub_id"), ("Teil", "asset_sub_id"), ("Lfd.", None),
    ("Unterpos.", "asset_sub_id"), ("SNr", "asset_sub_id"), ("Unter-nummer", "asset_sub_id"), ("Anlage", "asset_id"),
    # Short spellings are exact tokens, never substrings of other headers.
    ("lfd Buchwert", None), ("Teilwert", None), ("Subvention", None), ("Anlagengruppe", None),
])
def test_sub_number_and_asset_spellings_are_exact(header, expected):
    from carthagent.structure import IDENTITIES
    assert match(header, IDENTITIES) == expected
