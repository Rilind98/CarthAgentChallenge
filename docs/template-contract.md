# Arbeitsversion layout 20 contract

The workbook supplied during implementation has `A3_LAYOUT_VERSION = 20` and
SHA-256 `0449e028926cbe6e32c32fabc6da27a2693c36147c581c8d3b008e9543239668`.
Its 27 worksheets include both target schedules and other audit workflows.
The application fills supported fixed-asset schedules and preserves unrelated
worksheets; embedded descriptions are schema evidence, not authorization to
expand the application into invoice, trial-balance, or financial-statement work.

The private template belongs in `hackathon-local/templates/`, excluded from Git.
An operator supplies its path explicitly. The official JSON schema and scorer
have still not been supplied; using the workbook does not establish leaderboard
compatibility.

## Target worksheets and row boundaries

| Worksheet | Header | Data | Reserved regions |
|---|---|---|---|
| Anlagenverzeichnis | Row 1 | Rows 2–501 | Controls 503–508; hidden row ID column AQ |
| Gitter nach Anlagenklasse | Groups 5; headers 6; units 7 | Rows 8–27 | Printed-total placeholder 28; notes 30–31; hidden row ID X |
| Gitter nach Bilanzposition | Groups 5; headers 6; units 7 | Rows 8–27 | Printed-total placeholder 28; notes 30–31; hidden row ID X |

The two grids merge D5:I5, J5:R5, and S5:T5. Register data cells are unmerged.
Fiscal year-end in grid C2 drives the existing date-header formulas. A year alone
does not establish a fiscal closing date. Preserve headers unless source evidence
or an explicit run setting supplies that date.

An exporter must either safely expand data regions and all affected references,
or report overflow as unresolved projection. It must never overwrite controls,
silently drop facts, or claim complete output after truncation.

## Amount mapping

Stable field IDs come from `Feldkatalog Gitter` and
`Feldkatalog Anlagenverzeichnis`, not guessed display-header positions.

| Canonical source concept | Template field ID | Register | Grid |
|---|---|---|---|
| Opening acquisition cost | ahk_01_01 | W | D |
| Additions | ahk_zugang | X | E |
| Disposals at cost | ahk_abgang | Y | F |
| Cost transfers | ahk_umbuchung | Z | G |
| Net cost movement | ahk_bewegung_netto | No direct field | H |
| Closing acquisition cost | ahk_31_12 | AA | I |
| Opening accumulated depreciation | afa_01_01 | AB | J |
| Planned depreciation | afa_gj_plan | AC | K |
| Total current-year depreciation | afa_gj | No direct field | L |
| Extraordinary depreciation | afa_gj_ausserplan | AD | M |
| Depreciation on disposals | afa_abgang | AE | N |
| Reversals | zuschreibung | AF | O |
| Depreciation transfers | afa_umbuchung | AG | P |
| Net depreciation movement | afa_bewegung_netto | No direct field | Q |
| Closing accumulated depreciation | afa_31_12 | AH | R |
| Current closing book value | bw_31_12_gj | AI | S |
| Previous closing book value | bw_31_12_vj | AJ | T |
| Capitalized borrowing costs | fk_zinsen | AK | U |

The total current-year depreciation cannot be placed in the planned column
without evidence that it is planned-only. The grid's extraordinary amount is
included in its total; it must not be added twice. Source concepts with no direct
destination remain in evidence JSON with an explicit unmapped-output issue.

## Identity and granularity

The register requires one physical asset per row, with A asset number and B
description. C/D hold client class name/number; F/G client balance-position
number/name; H the account. Do not derive a balance-position label from an
account number alone. Subtotals do not belong in the register's data region.

Grid B contains the client's printed class/position number, C its name, V the
row type, and W a client balance-position name on the class grid only. Keep
source order and printed total/subtotal occurrences. Do not manufacture a serial
number when the client supplied none. A business context that cannot fit the
target's identity must produce a projection conflict, not deduplication.

## Representation and formulas

Amounts are expressed in EUR. Record every scale/polarity transformation in the
output binding while leaving raw source facts unchanged. The template describes
ordinary movement amounts as positive and explicitly describes transfer and net
movement fields as signed. Its register catalog requires positive disposals and
accumulated depreciation, including conversion from SAP negative credit displays.
Apply field-specific rules, never blanket absolute-value normalization.

Source printed closing balances override the corresponding AA/AH/AI/AJ template
formulas. Those source values can be inconsistent with the formula; preserve and
flag them. If a printed source value is unreadable or ambiguously mapped, the
destination must remain blank: an existing formula must not silently replace
the abstention.

The catalog permits some formulas when a value is genuinely absent from the
source. Such a value is a **derived output**, not a validated extracted fact.
It needs separate lineage and complete prerequisites, and must not receive a
fabricated source pointer. If derivation is not enabled/validated, clear the
active-row formula and record the missing value instead. Preserve unrelated
template formulas and workbook features.

## Open points and verification

`Offene Punkte` has column headers in row 3: Reiter, Stelle, Frage, Stand.
Replace the initial `Keine offenen Punkte` marker when any issue is present.
Include ambiguous source values, unsupported fields, template projection
conflicts, capacity limits, unknown period/context, and missing JSON/scorer
contract details as appropriate. Do not place all issues in a separate hidden
file while the workbook says there are none.

Acceptance checks compare sheet names/order, merges, validations, names, hidden
dimensions, and unaffected formulas/styles to the original. Every written value
must match its output binding and the common JSON snapshot. Formula calculations
are never read back as if they were source observations. No source template is
modified in place.
