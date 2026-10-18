# Book3 — layout profile

- **doc_type per sheet**: Sheet1 is a Zugangsliste (additions list for one fiscal year, German tax valuation area, grouped by balance-sheet account).
- **sheets**: 1 visible tab "Sheet1". Used range A1:AK162 (37 cols). No hidden sheets.
- **export origin guess**: German ERP fixed-asset module (non-SAP), a print report exported cell-for-cell to Excel. Every value sits at its print x-position, so it uses a wide, sparse grid of 37 narrow columns. The font is Arimo 8 (a metric clone of Arial), which points to a server-side or PDF-style conversion.
- **years/periods**: in the title block: row 1 has the report title "Zugänge Bilanzkonto - <valuation area> - bis <period>". Row 2 (right side) has "Periode MM/YYYY bis MM/YYYY". The class footers have "AHK (MM/YYYY):" labels.
- **header block above the table**: rows 1–2 are the report title, the valuation area ("Bestandsführung: …"), the grouping key ("Verdichtung: Bilanzkonto - alle Anlagegüter") and a run timestamp "erstellt am: dd.mm.yyyy um: hh:mm:ss". All are bold text (title 12 pt, rest 9 pt). Rows 3–7 hold **5 heading rows** that match the 3-line record layout below. There are no merged cells.
- **columns** (heading text verbatim, by heading row):
  - row 3: Bilanzkonto
  - row 4: Anlagen-Hnr. | Hauptbezeichnung
  - row 5: Unternr. | Bezeichnung | Methode | Afa-Beginn | ND (jj-mm) | Schicht-KZ | Afa% | Bilanzkonto | Kostenstelle | Kostenträger
  - row 6: Standort | Lieferant (Unr.)
  - row 7: Bucher | erfasst am | Belegdatum | Belegnr. | F-Belegnr. | Belegtext | Bezugsdatum | Bewegung | Betrag | Buchungstext (Pos) | Personenkonto (Beleg)
  - No hidden columns. Most columns are only spacers. The amount sits in column Y, one column left of its heading "Betrag" (Z).
- **row levels** (no outline, no indent, levels shown by row type and position):
  1. Balance-sheet account group header: "<6-digit account> <account name>" in column A, bold.
  2. Main asset: number (text) in B, main description in G.
  3. Sub-number record, 3 lines: (a) sub-no. (number), description, method, AfA start, useful life, shift code, rate, account, cost centre, cost unit; (b) location, supplier; (c) user, entry date, document date, document no., reference date, movement code, amount (bold), posting text.
- **subtotals/totals**
  - A main asset with ≥2 sub-numbers gets "Gesamtsumme Anlagenhauptnummer <no.>" in B, then a row with "Zugänge:" + amount. A single sub-number gets no subtotal.
  - Account group footer (one row): "<account> <name>" plus label/value pairs on the same row: "AHK (MM/YYYY):" <start> | "Zugänge:" | "SoPo:" | "Umbuch. AHK:" | "Zuschuss:" | "AHK (MM/YYYY):" <end>.
  - Grand total row "Gesamtsumme" with the same label/value pairs. A blank row sits before it.
  - All totals are **values**. There are no formulas in the workbook.
- **number storage**: amounts are real numbers, format "#,##0.00", bold. The sub-number is a number with format `#,##0;\-#,##0;0`. Everything else is **text with format "@"**: asset numbers, account numbers, document numbers, movement code, AfA rate ("d,dd %" with a German comma), useful life. Additions are positive. SoPo, Umbuchung and Zuschuss are written as 0. Units are EUR implied; no currency column.
- **dates / useful life / method codes**: all dates are **text** "dd.mm.yyyy". Useful life is text "jj-m" (years-months, e.g. "<y>-0"). Method is "<digit> <text>" (straight-line, or immediate GWG write-off). Shift code is "-". Movement code is a 4-digit text code; the posting text is the plain word for a new addition.
- **asset id pattern**: main asset number is 3 digits as text. The sub-number is a small integer (1, 2, 3 …). The account is a 6-digit text with a leading zero kept, followed by its name in the same cell.
- **arithmetic**: sum of the record amounts = "Zugänge" in each group footer and in the grand total (ties). AHK end = AHK start + Zugänge (the start is 0 because the list holds only new assets).
- **quirks**: a 3-line record whose fields are spread over a wide grid. The amount column is offset from its heading. Label/value pairs share one footer row. Supplier and cost-centre cells combine "<code> <name>". Text is cut off at a fixed print width. No fills, no freeze panes, no filters, no comments.
- **size**: 162 rows, 4 account groups, ~20 main assets, ~35 sub-number records.
- **generator notes**
  - Spread each record over 3 rows on a 37-column sparse grid, with headings in 5 rows (3–7) at the same x-positions.
  - Store dates, rates, useful life and ids as text (dd.mm.yyyy, "d,dd %", "jj-m"). Store only amounts (bold, "#,##0.00") and the sub-number as numbers.
  - Emit a "Gesamtsumme Anlagenhauptnummer" + "Zugänge:" pair only for main assets with ≥2 sub-numbers.
  - Put the group footer and the grand total as label/value pairs on one row. Write values only, no formulas.
  - Use Arimo/Arial 8 pt, a bold title block with the run timestamp and the period on the right.
