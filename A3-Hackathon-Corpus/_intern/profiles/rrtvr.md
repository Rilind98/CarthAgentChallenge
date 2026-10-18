# rrtvr.pdf — layout profile

- **doc_type**: Anlagenspiegel/Anlagengitter at single-asset level (SAP FI-AA asset history sheet, "Anlagengitter auf Basis gebuchter AfA", depreciation area 01 Handelsrecht, wide 13-column HGB §268(2) version).
  - p.1: report selection/statistics cover page (3 small boxed tables: Selektionsfeld/Wert, Sortierkriterien, Datenstatistik).
  - p.2–59: the asset grid, one page break per Anlagenklasse.
- **format**: no text layer, but NOT a raster scan. Printed from SAP to PDF, then sent through a PDF printer (Ghostscript) that turned every glyph into vector outlines. Pure black on white, no noise, no skew, no grey, no stamps, no handwriting, no blank pages. Text cannot be selected, so OCR is needed. Every page is portrait, with the landscape report rotated 90° counter-clockwise on it (text reads bottom to top).
- **pages / orientation / size / density**: 59 pages. Page box 595 x 791 pt (A4 width, US-letter height), portrait, content rotated. Small monospace (Courier-like) font of about 6–7 pt in landscape. Dense: 16 columns over the full width, about 45–50 rows on a full page.
- **software family**: SAP ECC/S4 FI-AA (asset history sheet report, ALV-grid list output with boxed cells).
- **years/periods**: one fiscal year. The period appears only in the report header (Berichtsdatum = fiscal-year end, Erstellungsdatum = print date a few months later). Column headings carry a period suffix: "GJ-Beg" (start of fiscal year), "des Jahres", "lfd", "aktuelle", "kumulierte".
- **header block / footer**:
  - Every page has a 2-line report header top left: "Berichtsdatum: dd.mm.yyyy" + report title, then "Erstellungsdatum: dd.mm.yyyy" + legal/version subtitle. The page number (bare integer) is top right. A full-width rule sits under the header.
  - Under the rule is a group header block: a label line "Buchungskreis | GeschBereich | Bilanzposition | BestandskontoAHK | Anlagenklasse" and a value line under it (GeschBereich empty).
  - No footer, no "page x of y", no company name on grid pages.
  - Cover page selection lines: Buchungskreis, Berichtsdatum, Bewertungsbereich 1, Sortiervariante, Anlageneinzelliste X, ALV Grid Control benutzen X, Gitterversion, Gebuchte Abschreibung X.
  - Cover page sort table (Aufst./Abst./Zwischensumme): Bilanzposition, Bestandskonto AHK, Anlagenklasse. Only Anlagenklasse has a subtotal.
  - Cover page statistics: "Übergebene Sätze", "Ermittelte Summenzeilen".
- **columns** (1 heading row, no spanned headings, every cell boxed):
  1. Anlage
  2. UNr.
  3. Anlagenbezeichnung
  4. Währg
  5. Aktivdatum
  6. AHK GJ-Beg
  7. AfA GJ-Beg
  8. Buchwert GJ-Beg
  9. Zugang
  10. AfA des Jahres
  11. Abgang
  12. AfA Abgang
  13. Umbuchung
  14. lfd Buchwert
  15. aktuelle AHK
  16. kumulierte AfA
- **row levels**: Buchungskreis > Bilanzposition (7 digits) > BestandskontoAHK (5–6 digits) > Anlagenklasse (3–4 digits) > Anlage > Unternummer.
  - The upper levels appear only in the group header block above the table. They are never shown as rows inside the table.
  - Asset rows are flat, with no indent and no bold.
- **subtotals/totals**:
  - One subtotal per class, in its own boxed row after the last asset: "*Klasse <class code> <class short name>". The label spans the Anlage to Anlagenbezeichnung columns, Währg is filled, and Aktivdatum is empty.
  - No Bestandskonto, Bilanzposition or grand total. The file ends on the last class subtotal.
  - No Übertrag / carry-over lines. A class that runs over a page break simply continues on the next page under a repeated header.
- **number format**: "." thousands, "," decimal, always 2 decimals. Negatives have a trailing minus ("1.234,56-"). Zero is shown as "0,00", never blank. Amounts are in EUR (Währg column), not TEUR. Numbers are right-aligned.
- **date / useful life / method**: Aktivdatum dd.mm.yyyy. No useful-life column and no depreciation-method key.
- **asset id**: plain integer of 1–4 digits, no leading zeros, roughly ascending by acquisition within a class. UNr. is almost always "0".
- **arithmetic**:
  - Accumulated depreciation is negative and cost is positive.
  - aktuelle AHK = AHK GJ-Beg + Zugang + Abgang + Umbuchung (Abgang negative).
  - kumulierte AfA = AfA GJ-Beg + AfA des Jahres + AfA Abgang.
  - lfd Buchwert = aktuelle AHK + kumulierte AfA.
  - Buchwert GJ-Beg = AHK GJ-Beg + AfA GJ-Beg.
  - Disposals are shown at cost: Abgang = minus the full AHK, AfA Abgang = plus the accumulated depreciation reversed, and the remaining book value goes to 0,00.
  - Many fully depreciated assets keep a 1,00 Erinnerungswert.
  - Current-year acquisitions show 0,00 in all GJ-Beg columns.
  - Umbuchung holds capitalisations out of assets under construction.
- **quirks**:
  - Fixed-width description column: long descriptions are cut off mid-word, not wrapped.
  - Header and column heading row repeat on every page.
  - A new page at every class change gives many short pages (1–3 rows plus subtotal) and large white areas.
  - The cover page is separate, with sort and statistics tables.
  - Report subtitle marks the grid as the wide but incomplete 13-column version.
  - Descriptions sometimes start with an internal tool/part number of the form "NNN-NNNNNN".
  - Same ledger and year as zt.pdf (class summary) and zunzun.pdf (additions list). Class totals and Zugang amounts tie across the three files.
- **size**: about 2,000 asset rows, about 28 classes/subtotals, about 8 Bilanzpositionen.
- **generator notes**:
  - Render a boxed ALV grid in monospace with the 16 columns above, landscape. Then put it on a portrait page rotated 90° CCW and convert the text to outlines (no text layer, but crisp vector, zero raster noise).
  - Add the 2-line SAP report header plus page number and the 5-field group header block on every page. Force a page break at every class change.
  - Close each class with only a "*Klasse <code> <name>" subtotal row. No higher totals and no grand total.
  - Use trailing-minus negatives, explicit 0,00 and many 1,00 Erinnerungswert rows. Show disposals at cost, with AfA Abgang reversal and book value 0,00.
  - Cut descriptions at a fixed character width. Add a cover page with the Selektionsfeld/Sortierkriterien/Datenstatistik tables.
  - For a full set, generate the matching class summary (zt-like) and additions list (zunzun-like) from the same synthetic ledger so the totals tie.
