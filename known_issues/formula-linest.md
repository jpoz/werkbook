# formula/LINEST

- Audit report: `../testdata/out/issues/formula/LINEST.json`
- Diff count: `1`
- Category: `statistical function`
- Summary: A `LINEST(...,TRUE)` result path diverges from Excel and currently returns `#NUM!` instead of the expected numeric output.

Representative mismatches:

- `Results!B26`: `INDEX(LINEST(Data!E1:E5,Data!F1:F5,,TRUE),4,1)` -> Excel `1.1589948344943812E+32`, werkbook `error:#NUM!`
