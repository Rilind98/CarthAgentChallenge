# Book5 — layout profile

- **doc_type per sheet**: Sheet1 is a SuSa (trial balance: opening balance + cumulative debit/credit, all accounts including P&L).
- **sheets**: 1 tab "Sheet1", A1:E273 (formatted blank rows continue to ~288). No hidden sheets.
- **export origin guess**: accounting program account list exported to Excel, re-saved in recent Excel (Aptos Narrow 11).
- **years/periods**: not printed anywhere.
- **header block above the table**: none. Row 1 is the heading.
- **columns**: Konto (bold) | Kurztext (bold) | "  Saldovortrag S/H" | " Kumuliert Soll" | " Kumuliert Haben". Some headings have leading spaces. 1 heading row. No closing-balance column.
- **row levels**: one flat level, sorted by account ascending. No groups.
- **subtotals/totals**: none.
- **number storage**: numbers as numbers, "#,##0.00". The opening balance is signed (credit negative). Debit and credit are positive. Zeros are written as 0.
- **dates / useful life / method codes**: none.
- **asset id pattern**: account is a 5–6 digit integer. The fixed-asset accounts come first. Kurztext is cut to 20 characters.
- **arithmetic**: not checked (no totals or closing balance in the file).
- **quirks**: leading spaces in headings. Truncated 20-character account names. No total row.
- **size**: ~270 accounts.
- **generator notes**
  - 5 columns with no header block, accounts ascending, 20-character short names.
  - Signed opening balance plus cumulative debit and credit. No closing balance, no totals.
