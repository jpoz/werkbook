# apache-poi/NumberFormatApproxTests

- Audit report: `../testdata/out/issues/apache-poi/NumberFormatApproxTests.json`
- Diff count: `13`
- Category: `error typing`
- Summary: `TEXT(...)` cases that Excel surfaces as an error cell are coming back from werkbook as the string literal `#VALUE!`.

Representative mismatches:

- `Tests!A2`: `TEXT(C2, B2)` -> Excel `error:#VALUE!`, werkbook `string:#VALUE!`
- `Tests!A3`: `TEXT(C3, B3)` -> Excel `error:#VALUE!`, werkbook `string:#VALUE!`
- `Tests!A4`: `TEXT(C4, B4)` -> Excel `error:#VALUE!`, werkbook `string:#VALUE!`
- `Tests!A21`: `TEXT(C21, B21)` -> Excel `error:#VALUE!`, werkbook `string:#VALUE!`
