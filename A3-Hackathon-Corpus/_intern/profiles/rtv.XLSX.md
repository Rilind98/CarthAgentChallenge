# Profile: rtv.XLSX

- **doc_type:** Anlagenspiegel/Anlagengitter (HGB, class level, no single assets). Report title "Anlagengitter - 01 Handelsrecht", subtitle "HGB Par.268,2 13-Spalten (breite Version) (unvollständig)".
- **format:** text-PDF made from an xlsx: an SAP ALV list exported to Excel, then printed to PDF with a Ghostscript-based PDF printer. Excel cell borders and fills survive.
- **pages / orientation / size / feel:** 5 pages, landscape, A4. Courier New 6.8 pt (regular + bold), monospace, very dense. The table uses only the left ~60% of the page width; the rest is blank.
- **software family guess:** SAP-like (FI-AA asset history sheet, ALV grid output with star-prefixed subtotal levels and a sort/statistics cover page).
- **years/periods:** one fiscal year, reported in-year: header "Berichtsdatum: DD.MM.YYYY" (a date inside the year, not year end, hence "(unvollständig)") and "Erstellungsdatum: DD.MM.YYYY" (run date, later). Year only implicit in the "GJ-Beg" (fiscal-year start) column names.
- **header block:** page 1 is a cover page with two small boxed tables: "Sortierkriterien | Aufst. | Abst. | Zwischensumme" (rows Buchungskreis, Geschäftsbereich, Bilanzposition, Bestandskonto AHK, Anlagenklasse, marked with "X") and "Datenstatistik | Anzahl" (rows "Übergebene Sätze", "Ermittelte Summenzeilen"); header cells blue. Pages 2-5: two header lines, labels bold blue ("Berichtsdatum:", "Erstellungsdatum:") followed by report name and variant text; page number alone at far right of line 2; horizontal rule below. No client name, no "of y".
- **footer:** none.
- **columns (in order):** 4 heading rows in one blue band, repeated on every page:
  1. BuKr | GsBe | BilanzPos | BestKtoAHK | Klasse | Währg
  2. AHK GJ-Beg | Zugang | Abgang | Umbuchung | Nachaktivg | Invest.Förderg | aktuelle AHK
  3. AfA GJ-Beg | AfA des Jahres | AfA Abgang | AfA Umbuchung | AfA Nachaktivg | Zuschreibungen | kumulierte AfA
  4. Buchwrt GJ-Beg | (blank x5) | lfd Buchwert
  So 13 value columns stacked into 3 lines of 7/7/2 values. No spanned headings.
- **row levels:** Buchungskreis > Geschäftsbereich > Bilanzposition > Bestandskonto AHK > Anlagenklasse. Every record is a 4-line block inside a bordered box: key line (BuKr 4 digits, GsBe usually blank, BilanzPos 7 digits, BestKtoAHK 4-5 digits, Klasse "L"+4 digits, class name in a lilac-shaded cell, "EUR"), then AHK line, AfA line, Buchwert line. Level shown only by a star prefix in column 1: none = Anlagenklasse, "*" = Bestandskonto subtotal, "**" = Bilanzposition subtotal, "***" = Geschäftsbereich subtotal, "****" = Buchungskreis total (glued: "****NNNN"). Subtotal blocks are filled pale yellow; some class blocks are bold on a grey band.
- **subtotals/totals:** no words like "Summe" at all; star prefix only. A subtotal is printed even when it equals the single class above it (so most classes appear 2-3 times). Grand total = "****" block on the last page, preceded by "***".
- **number format:** "." thousands, "," decimals, 2 decimals, trailing minus for negatives ("1.234,56-"). Accumulated depreciation carried negative; Abgang (cost) negative, AfA Abgang positive. Zero always printed "0,00" (never blank). Currency per block "EUR".
- **date format:** DD.MM.YYYY in header only. No useful life, no method, no dates in body.
- **asset id pattern:** none (no single assets). Class keys "L" + 4 digits; Bestandskonto 4-5 digits; Bilanzposition 7 digits.
- **arithmetic:** verified on all 35 blocks: AHK GJ-Beg + Zugang + Abgang + Umbuchung + Nachaktivg + Invest.Förderg = aktuelle AHK; AfA GJ-Beg + AfA des Jahres + AfA Abgang + AfA Umbuchung + AfA Nachaktivg + Zuschreibungen = kumulierte AfA; AHK GJ-Beg + AfA GJ-Beg = Buchwrt GJ-Beg; aktuelle AHK + kumulierte AfA = lfd Buchwert. Disposals at cost (Abgang) with the matching AfA Abgang. Subtotals and grand total tie. Cover-page counts equal the number of class blocks and of sum blocks.
- **quirks:** right-most column clipped at the Excel page edge (headings cut to "aktuelle AH", "kumulierte AF", "lfd Buchwer"; trailing minus partly cut); stacked 3-line value layout per record; star-only level markers; all-zero classes printed (e.g. down-payment / construction classes); cover page with sort criteria; interim report date; landscape page mostly empty on the right.
- **size:** 0 asset rows; ~15 asset-class blocks plus ~20 subtotal blocks; 1 company code, 1 business area, ~8-10 balance-sheet positions.
- **generator notes:**
  - Emit an SAP-ALV-in-Excel look: Courier New, blue header band, lilac class-name cell, pale-yellow subtotal blocks, boxed 4-line records.
  - Each record: key line + 3 value lines (7 / 7 / 2 values) in the 13-column HGB layout; enforce the four row laws exactly.
  - Use star prefixes (*, **, ***, ****) for subtotal levels and print redundant subtotals even for single-child groups.
  - Trailing-minus negatives, "0,00" for zero, depreciation as negatives.
  - Add the cover page (Sortierkriterien + Datenstatistik) and the two-line Berichtsdatum/Erstellungsdatum header with an in-year report date and "(unvollständig)".
  - Clip the last column at the page edge to test readers on cut values.
