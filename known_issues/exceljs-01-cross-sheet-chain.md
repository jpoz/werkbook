# exceljs/01_cross_sheet_chain

- Audit report: `../testdata/out/issues/exceljs/01_cross_sheet_chain.json`
- Diff count: `30`
- Category: `recalc / dependency`
- Summary: Cross-sheet chained formulas in `Model` are staying at `0` instead of propagating the values from `Inputs` through the chain.

Representative mismatches:

- `Model!B2`: `Inputs!B2*(1+Inputs!$H$2)` -> Excel `126`, werkbook `0`
- `Model!C2`: `B2*(1+Inputs!$H$2)` -> Excel `132.3`, werkbook `0`
- `Model!D2`: `C2*(1+Inputs!$H$2)` -> Excel `138.91500000000002`, werkbook `0`
- `Model!E2`: `D2*(1+Inputs!$H$2)` -> Excel `145.86075000000002`, werkbook `0`
