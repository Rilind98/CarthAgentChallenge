# zt.pdf — layout profile

- **doc_type**: Anlagenspiegel/Anlagengitter, summary level. This is the SAP FI-AA depreciation report "Abschreibungen - 01 Handelsrecht" (display variant for depreciation), totalled per Anlagenklasse with Bilanzposition subtotals. No single assets.
  - p.1: selection/statistics cover page.
  - p.2–3: class summary table.
- **format**: text-PDF with a real text layer. Courier 8.3 pt, vector boxes, no scan effects. Fully clean.
- **pages / orientation / size / density**: 3 pages, landscape 791 x 595 pt (US-letter-height x A4-width box). Medium density: about 40 table lines on p.2. p.3 holds only an orphaned last subtotal.
- **software family**: SAP ECC/S4 FI-AA ALV list (asset depreciation report), printed straight to PDF.
- **years/periods**: one fiscal year. The period appears only in the header: "Berichtsdatum: dd.mm.yyyy" (fiscal-year end) and "Erstellungsdatum: dd.mm.yyyy" (print date). Column headings carry no dates. The pair "Ende-Buchwert / Beg-Buchwert" gives closing and opening values.
- **header block / footer**:
  - Every page has a 2-line header top left (Berichtsdatum + report title, Erstellungsdatum), the page number top right, and a rule under the header.
  - No footer and no company name.
  - Cover page, Selektionsfeld/Wert: Buchungskreis, Berichtsdatum, Bewertungsbereich 1, Sortiervariante (alphanumeric Z-variant), a flag line "X", Anzeigevariante "/<name>", ALV Grid Control benutzen X.
  - Cover page, Sortierkriterien (Aufst./Abst./Zwischensumme): Bilanzposition (with subtotal), Bestandskonto AHK, Anlagenklasse.
  - Cover page, Datenstatistik: Übergebene Sätze, Ermittelte Summenzeilen.
- **columns** (1 heading row, no spanned headings, boxed table, column heading repeats per page):
  1. BilPos
  2. Bezeichnung
  3. Klasse
  4. kum AnschWert
  5. Bew AnschWert
  6. kum AfA
  7. gep AfA
  8. Bew Wertber
  9. Ende-Buchwert
  10. Beg-Buchwert
  11. Währg
- **row levels**: Bilanzposition > Anlagenklasse.
  - Each class row starts with the 7-digit Bilanzposition code, then the short class name and the class code (3–4 digits).
  - Class rows of one Bilanzposition form one boxed group.
- **subtotals/totals**:
  - After each group comes a boxed subtotal row "*<7-digit Bilanzposition>", with an empty description and the amounts.
  - No grand total. The file ends on the last Bilanzposition subtotal, which spills alone onto the last page.
- **number format**: "." thousands, "," decimal, 2 decimals, trailing minus for negatives. Zero is shown as "0,00". Währg column is "EUR", in units, not TEUR.
- **date / useful life / method**: only header dates, dd.mm.yyyy. No useful life and no method.
- **asset id**: none (class level only). Class code 3–4 digits, Bilanzposition 7 digits.
- **arithmetic**:
  - kum AnschWert = opening cost.
  - Bew AnschWert = net cost movement of the year (additions, disposals, transfers; can be negative).
  - kum AfA = opening accumulated depreciation (negative).
  - gep AfA = planned depreciation of the year (negative).
  - Bew Wertber = depreciation movement on disposals/transfers.
  - Ende-Buchwert = kum AnschWert + Bew AnschWert + kum AfA + gep AfA + Bew Wertber.
  - Beg-Buchwert = kum AnschWert + kum AfA.
  - Disposals are netted at cost inside Bew AnschWert, with their depreciation in Bew Wertber.
- **quirks**:
  - Columns are compressed into movement pairs (Bew = movement) instead of the full Zugang/Abgang/Umbuchung split.
  - Abbreviated class names cut to about 20 characters.
  - An orphan subtotal sits alone on the last page.
  - One class row appears with all zeros.
  - Has a cover page.
  - This is the class-level twin of rrtvr.pdf: the same classes, and the totals tie.
- **size**: about 29 class rows, about 8 Bilanzposition subtotals.
- **generator notes**:
  - Emit as a real text PDF in Courier about 8 pt, landscape, with an SAP 2-line header plus page number and a boxed ALV table.
  - One row per class, prefixed with the Bilanzposition code. "*<BilPos>" subtotal after each group, no grand total.
  - Use the 11 columns and the row law above. Ende-Buchwert must reconcile across the 5 value columns.
  - Use trailing-minus negatives and explicit 0,00. Include an all-zero class and let the last subtotal break onto its own page.
  - Add the selection/statistics cover page. Derive the values from the same synthetic asset ledger as the rrtvr-like grid so the class totals match.
