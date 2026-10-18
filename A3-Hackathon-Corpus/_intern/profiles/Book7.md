# Book7 — layout profile

- **doc_type per sheet**: Sheet1 is a SuSa (trial balance: opening balance + annual debit/credit, all accounts including P&L and carry-forward accounts).
- **sheets**: 1 tab "Sheet1", A1:E288. No hidden sheets.
- **export origin guess**: account-list export from a DATEV/SKR03-style accounting program (3–4 digit accounts, "Saldenvorträge" accounts at the end), re-saved in recent Excel.
- **years/periods**: not printed anywhere.
- **header block above the table**: none. Row 1 is the heading (fill #B4B4B4, text format "@").
- **columns**: Konto | Beschriftung | EB-Wert S/H | JVZ Soll | " JVZ Haben" (with a leading space). 1 heading row.
- **row levels**: one flat level, accounts ascending. No groups.
- **subtotals/totals**: none.
- **number storage**: numbers as numbers. The opening balance column is **bold**, "#,##0.00", signed (credit negative). Debit and credit use the long format `##,###,###,###,###,###,##0.00`. Debit and credit are **blank** (not 0) on accounts without movements. The account is an integer with format "#####", so the SKR leading zero is lost (a 4-digit "0nnn" account shows as "nnn").
- **dates / useful life / method codes**: none.
- **asset id pattern**: account is a 3–4 digit integer. The leading zero is dropped. Names are up to 40 characters.
- **arithmetic**: not checked (no totals or closing balance).
- **quirks**: bold opening-balance column. Lost leading zeros. Blank debit/credit cells on some rows. Font is 12 pt in the data and 11 pt in the opening-balance column.
- **size**: ~290 accounts.
- **generator notes**
  - 5 columns with a grey heading. SKR03-style accounts stored as integers without the leading zero.
  - Bold signed opening balance. Leave debit/credit blank for accounts without movements. No totals.
