# Profile: asdasd

- doc_type: Anlagenspiegel/Anlagengitter, gross form with full asset detail (one line per asset, AHK history + movements + cumulative AfA + book values). Sister print of aasdaswd (same run, same position/group tree).
- format: text-PDF (Arial, real text layer); printed from an accounting program (fixed-asset module), not Excel. Re-saved through macOS print after redaction.
- pages: 39. orientation: landscape. page size: A4. font feel: medium-dense (7 pt; asset rows regular, headings/sums bold; ~15 asset rows per page because many descriptions wrap to 2 lines).
- software family guess: German mid-market FiBu / Anlagenbuchhaltung (Sage-Classic-Line-like report style).
- years/periods: one fiscal year. Period printed in the title ("Anlagespiegel 31.12.YYYY steuerrechtlich", valuation area = tax) and in the last column heading "Buchwert 31.12."; previous year only as "Buchwert Vorjahr".
- header/footer: title bold ~10 pt top left; redaction box top right (company block). Column heading block repeated on every page. Footer: "Datum:DD.MM.YY" (no space after colon) bottom left, "Seite: n" bottom right (no "of y"). No client number, no selection lines.
- columns (4 heading lines + 1 unit line; no spanned headings):
  1. Inventar-Nummer
  2. Inventar-Bezeichnung
  3. Anschaff. Datum
  4. AfA Art
  5. ND
  6. Anschaff.-/Herstellungskosten (hist.) — EUR
  7. Zugänge, Geschäftsjahr — (+)EUR
  8. davon aktivierte Zinsen Fremdkapital — (+)EUR
  9. Abgänge / Minderungen Geschäftsjahr — (-)EUR
  10. Umbuchungen Geschäftsjahr — (+/-)EUR
  11. Zuschreibungen Geschäftsjahr — (+)EUR
  12. Abschreibungen Geschäftsjahr — EUR
  13. Buchwert Vorjahr — EUR
  14. Abschreibungen (kumuliert) — (-)EUR
  15. Buchwert 31.12. — EUR
- row levels: Bilanzposition > Gruppe > Konto > Anlage.
  - Bilanzposition: flush left, bold, numbered "0." to "9.", client-defined labels (not strict HGB §266 order, no Finanzanlagen).
  - Gruppe: indent 1, bold, name only.
  - Konto: indent 2, bold, "<account no> <name>"; sometimes missing (assets hang directly under the group).
  - Anlage: inventory number right-aligned in a narrow column, description left-aligned right next to it (text extraction glues them: "<id><Description>"), then date, method code, ND, figures.
- subtotals/totals: "Summe <account no> <account name>" (indent 2), "Summe <group name>" (indent 1), "Summe <n>. <position name>" (flush left); all bold, all columns filled. No grand total: file ends after the last position sum.
- number format: thousands ".", decimal ",", two decimals; sign carried by the column (unit row), no minus signs. On asset rows empty movement columns are BLANK; AfA GJ, Buchwert Vorjahr, kum. AfA and Buchwert 31.12. always printed (as "0,00" when zero). On sum rows every column printed, zeros as "0,00". Unit EUR.
- date format: DD.MM.YY (two-digit year; acquisition years range over many decades). ND: integer years (0 to 50, 0 for land). AfA Art codes, lower case: "lin" (most), "d/l" (declining then linear), "geb" (building), "k" (none, land), rare "sb" and "hsv".
- asset id pattern: 1-4 digits, right-aligned, optional one-letter suffix for split items (lower case a-f series, sometimes upper case A); not tied to the account number.
- arithmetic: per asset AHK (hist.) − Abschreibungen (kumuliert) = Buchwert 31.12. and Buchwert Vorjahr − Abschreibungen GJ = Buchwert 31.12. (no movement). Full law: AHK + Zugänge − Abgänge ± Umbuchungen + Zuschreibungen − kum. AfA = Buchwert 31.12. Abgänge at historical cost (a disposed asset shows its AHK again in Abgänge and ends at 0,00, the prior one-euro book value just vanishes). About 6% of asset rows and ~40% of sum rows are off by 1-2 cents (rounding drift). Many fully written-down assets carry a one-euro memo value. No grand total to tie.
- quirks: descriptions wrap to a second line under the description column; asset id glued to description in the text layer; split assets with letter suffixes listed in a run; group without account level; same report tree as the net version; heading repeated per page, no carry-over lines; redaction box top right.
- size: about 570 asset rows, 10 balance-sheet positions, about 26 groups, about 26 accounts.
- generator notes:
  - Landscape A4, Arial 7 pt, 4-line heading + unit row with sign hints, heading on every page.
  - Four levels with three "Summe" levels; numbered positions from "0.", accounts sometimes skipped.
  - Asset rows: blank movement cells, always-filled AfA / BW columns, two-digit-year dates, lower-case method codes, integer ND.
  - Inventory numbers short (1-4 digits) with optional letter suffix runs; descriptions wrap.
  - Heavy share of one-euro memo values; disposals at cost; occasional 1-cent drift; no grand total.
