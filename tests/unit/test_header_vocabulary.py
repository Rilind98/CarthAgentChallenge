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


@pytest.mark.parametrize(("header", "expected"), [
    # R7-4: a depreciation word plus a movement word names the depreciation counterpart, never the AHK movement. Additions to
    # cumulative depreciation are the year's charge (the same convention as "Zugänge" under "Kumulierte Abschreibungen").
    ("AfA Zugänge", "total_afa_gj"), ("Abschreibungen Zugänge", "total_afa_gj"), ("Abschr. Zugang", "total_afa_gj"),
    ("AfA Abgänge", "afa_abgang"), ("AfA Abgang", "afa_abgang"), ("Abschreibungen Abgänge", "afa_abgang"),
    ("AfA Umbuchungen", "afa_umb"), ("Abschreibungen Umbuchung", "afa_umb"), ("AfA Zuschreibungen", "zuschreibung"),
    # Controls: plain movements and plain depreciation keep their meaning.
    ("Zugänge", "zugang"), ("Abgänge", "abgang"), ("Umbuchungen", "umb"), ("Abschreibungen des Geschäftsjahres", "total_afa_gj"),
])
def test_depreciation_movement_headers_name_the_depreciation_counterpart(header, expected):
    assert match(header, FIELDS) == expected


@pytest.mark.parametrize(("headers", "expected"), [
    (("Buchwert 31.12.2024",), {"Buchwert 31.12.2024": "bw_prev"}),
    (("Buchwert 31.12.2025",), {"Buchwert 31.12.2025": "bw_close"}),
    (("Buchwert 31.12.2025", "Buchwert 31.12.2024"), {"Buchwert 31.12.2025": "bw_close", "Buchwert 31.12.2024": "bw_prev"}),
])
def test_r7_4_a_dated_book_value_column_follows_the_reporting_year(tmp_path: Path, headers, expected):
    source = tmp_path / "bv.xlsx"
    workbook = Workbook(); sheet = workbook.active; sheet.title = "AV"
    sheet["A1"] = "Anlagenspiegel Geschäftsjahr 2025"; sheet["A2"] = "EUR"
    sheet.append(["Anlage", "Bezeichnung", *headers])
    sheet.append(["100001", "Maschine", *(1000 + index for index, _ in enumerate(headers))])
    workbook.save(source)
    observations = acquire_xlsx(source, "d", sha256(source.read_bytes()).hexdigest(), DEFAULT_BOUNDS).observations
    resolver = ContextResolver(observations); by_locator = {o.locator: o for o in observations}
    for index, header in enumerate(headers):
        context = resolver.resolve(by_locator[f"AV!{'CD'[index]}4"], observations)
        assert context.business is not None and (context.business.field, context.business.period) == (expected[header], "2025"), header


@pytest.mark.parametrize("header", ["SUMME AHK / Zugang / Abgang / Umbuchung", "Zugang/Umbuchung", "Zugänge und Abgänge"])
def test_r7_8_a_header_naming_several_movements_is_no_single_field(header):
    assert match(header, FIELDS) is None


@pytest.mark.parametrize(("header", "expected"), [
    # R8-1: the depreciation word must be the HEAD of the header; as an asset qualifier it never moves the column to AfA.
    ("Zugänge AfA-pflichtig", "zugang"), ("Zugänge (AfA-fähig)", "zugang"), ("Zugänge (ohne AfA)", "zugang"),
    ("Zugänge nicht abschreibbar", "zugang"), ("Zugänge (davon abschreibungsfähig)", "zugang"), ("Abschreibungsgruppe Zugang", "zugang"),
    ("Abgänge (nicht abschreibungspflichtig)", "abgang"), ("Zugänge keine AfA", "zugang"),
    # Head-noun forms keep the depreciation counterpart.
    ("AfA Zugänge", "total_afa_gj"), ("Abschreibungen Zugänge", "total_afa_gj"), ("Abschr. Abgänge", "afa_abgang"),
    ("Abschreibung auf Zugänge", "total_afa_gj"), ("AfA auf Abgänge", "afa_abgang"),
    # A combined write-up/depreciation (net) column names no field.
    ("Zuschreibungen/Abschreibungen", None), ("Abschreibungen / Zuschreibungen", None),
])
def test_r8_1_depreciation_word_counts_only_as_the_head(header, expected):
    assert match(header, FIELDS) == expected
