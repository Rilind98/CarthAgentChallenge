# Profile: aasdaswd

- doc_type: Anlagenspiegel/Anlagengitter, net form ("Netto-Anlagespiegel": starts from book value of prior year, no AHK column). Account-level, no individual assets.
- format: text-PDF (Arial, real text layer); printed from an accounting program (fixed-asset module), not Excel. Re-saved through macOS print after redaction.
- pages: 5. orientation: landscape. page size: A4. font feel: airy (7 pt, nearly all bold, generous row spacing, one blank half-line between blocks).
- software family guess: German mid-market FiBu / Anlagenbuchhaltung (Sage-Classic-Line-like report style). Sister print of file asdasd (same run, same layout family).
- years/periods: one fiscal year. Period printed in the title ("Netto-Anlagespiegel 31.12.YYYY steuerrechtlich", valuation area = tax) and in the last column heading "Buchwert 31.12.".
- header/footer: title bold ~10 pt top left; redaction box top right (company block). Column heading block repeated on every page. Footer: "Datum: DD.MM.YY" (print date) bottom left, "Seite: n" bottom right (no "of y"). No client number, no selection lines.
- columns (3-4 heading lines + 1 unit line; no spanned headings; first column heading "Bilanzposition"):
  1. Bilanzposition (text)
  2. Buchwert Vorjahr — EUR
  3. Zugänge, Geschäftsjahr — (+)EUR
  4. davon aktivierte Zinsen Fremdkapital — (+)EUR
  5. Abgänge / Minderungen Geschäftsjahr — (-)EUR
  6. Umbuchungen Geschäftsjahr — (+/-)EUR
  7. Zuschreibungen Geschäftsjahr — (+)EUR
  8. Normal-AfA Geschäftsjahr — EUR
  9. Sonder-AfA Geschäftsjahr — EUR
  10. Abschreibungen (kumuliert) — (-)EUR
  11. Buchwert 31.12. — EUR
  - The unit row carries the sign hint per column ("(+)EUR", "(-)EUR", "(+/-)EUR").
- row levels: Bilanzposition > Gruppe > Konto.
  - Bilanzposition: flush left, numbered "0." to "9." with client-defined labels (not strict HGB §266 order: intangibles numbered 0, then custom positions that split fixed assets by purpose; no Finanzanlagen).
  - Gruppe: indent 1, group name only (often equals the account name, sometimes with "Kto. <nr>" in the text).
  - Konto: indent 2, 3-4 digit account number + name, full row of figures.
- subtotals/totals: "Summe <group name>" after each group (indent 1); "Summe <n>. <position name>" after each position (flush left). Every group has a sum line even when it holds one account (sum line duplicates the account line). No grand total: file ends after the last position sum.
- number format: thousands ".", decimal ",", always two decimals; zero printed as "0,00" in every column; no negatives seen (sign carried by the column, see unit row). Unit EUR.
- date format: only in footer, DD.MM.YY. No useful life / method columns.
- asset id pattern: none (account level). Account numbers 3-4 digits, SKR04-like asset accounts, some with an extra trailing digit.
- arithmetic: Buchwert Vorjahr + Zugänge − Abgänge ± Umbuchungen + Zuschreibungen − Normal-AfA − Sonder-AfA = Buchwert 31.12. Holds on every line except about 1 in 7 sum lines off by 1-2 cents (rounding drift). Abgänge are at book value. "Abschreibungen (kumuliert)" is informative only, not part of the row law. Many accounts carry a one-euro memo book value.
- quirks: long labels wrap to a second line (including the account number alone wrapping onto line 2); a group can hold two accounts with different names but the sum line carries the first account's group name; no grand total; redaction box top right; everything bold.
- size: about 26 account rows, 10 balance-sheet positions, about 25 groups.
- generator notes:
  - Net Spiegel: start column "Buchwert Vorjahr", end column "Buchwert 31.12.", plus "Abschreibungen (kumuliert)" just before the end.
  - Unit row with sign hints "(+)EUR / (-)EUR / (+/-)EUR" under the headings.
  - Three levels with "Summe <name>" lines; numbered positions starting at "0.", custom labels.
  - Print every zero as "0,00"; inject occasional 1-cent rounding drift in sum lines.
  - Footer "Datum: DD.MM.YY" left and "Seite: n" right; heading repeated per page; no grand total.
