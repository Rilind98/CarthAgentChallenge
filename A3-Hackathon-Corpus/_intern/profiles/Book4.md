# Book4 — layout profile

- **doc_type per sheet**
  - Sheet1: other (empty sheet, A1 only).
  - Sheet2: Zugangsliste (additions list, commercial-law valuation area), pasted as plain text with one print line per cell.
- **sheets**: 2 visible tabs, default names. Sheet1 is empty. Sheet2 used range A1:A356 (**one column only**, width ~80). No hidden sheets.
- **export origin guess**: a print/PDF list from a German mid-market ERP accounting suite, copied as text into Excel. Every print line is one string in column A. There is no column structure. The page header/footer (report name, period, run number, product/version string, user, run timestamp, "Währung: EUR Seite N") repeats at each page break.
- **years/periods**: only inside the repeated page footer line: "Zugangsliste <area> dd.mm.yyyy bis dd.mm.yyyy". The record dates are inside the text lines.
- **header block above the table**: the file starts in the middle of the report. Rows 1–7 are the 7 column-heading lines. The title block of page 1 is missing. Each later page break inserts 3–4 lines (footer text glued to the end of the previous data line, then client number + user, run date/time, "Währung: EUR Seite N"), followed by the 7 heading lines again.
- **columns**: none in Excel terms. The print heading spans 7 lines, verbatim: "Bil.Pos. Anl.Konto Anlagenummer Lieferant" / "Bezeichnung" / "Standort Ansch-Dat." / "AfA-Meth. AfA-% ND" / "KSt KTr Buchungstext" / "Prj GE GF Anteil %" / "BA Belegdatum Menge Betrag". Fields inside a line are separated by single spaces.
- **row levels** (shown by text prefixes and bold, no indent, no outline):
  1. Balance-sheet item: shown only in its footer "* Summe Bil.Pos. <6 digits> <amount>" (bold).
  2. Asset account: header "<6-digit account> <name>" (bold). Footer "** Summe Anl.Konto <6 digits> <amount>" (bold).
  3. Asset: 4–6 lines: id + capitalisation date (+ optional supplier text) / description + method code + rate + useful life in months / share-percent line (full share) / movement line(s) / optional posting-text line. Footer "Summe <amount>" (bold).
  - The line order inside an asset block is not fixed (the share-percent line comes before or after the movement line).
- **subtotals/totals**: 3 levels of printed values (asset "Summe", account "** Summe Anl.Konto", balance-sheet item "* Summe Bil.Pos."). The last line "Gesamt" has a non-breaking space and no amount. The first account footer has no amount (the file starts mid-report). All are text; there are no formulas.
- **number storage**: **all numbers are text** in German format: thousands ".", decimal "," ("9.999.999,99"). The quantity has 4 decimals. Rate "dd,dd". Additions are positive. No minus signs seen.
- **dates / useful life / method codes**: dates are text "dd.mm.yyyy". Useful life is in **months** as an integer at the end of the description line. Method is a 3-letter code (straight-line). The movement type is a 2-letter code at the start of the movement line (new addition). A rate placeholder "******" can appear.
- **asset id pattern**: "NNNNNNN.NN" (7-digit main number + "." + 2-digit sub-number). Digits 2–3 of the number repeat the account (account "0AB000" gives ids "AB…"). Accounts are 6 digits with a leading zero.
- **arithmetic**: asset "Summe" = its movement amount(s). Account and balance-item sums tie wherever the lines are intact. Page breaks glue movement lines to footer text, so a naive line parser misses ~4 % of the asset sums and two account/balance-item checks.
- **quirks**: one-column text dump. Page header/footer text is glued into a data line at a page break. The report starts mid-stream (orphan "Summe" and account footer without a value at the top). Bold marks only the total and header lines. Arial 8. Sheet1 is an empty decoy.
- **size**: 356 lines, ~8 print pages, ~10 asset accounts, 3 balance-sheet items, ~50 assets.
- **generator notes**
  - Render a print list, then write **one line per cell in column A** with German-formatted numbers as text.
  - Repeat the page header and the 7-line column heading every ~55 lines, and glue the footer text onto the last data line of each page.
  - Use the 3-level "Summe" / "** Summe Anl.Konto" / "* Summe Bil.Pos." footers in bold and a final "Gesamt" with no value.
  - Vary the line order inside asset blocks, and start the file mid-report on some variants.
  - Add an empty first sheet.
