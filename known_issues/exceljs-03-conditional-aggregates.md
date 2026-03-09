# exceljs/03_conditional_aggregates

- Audit report: `../testdata/out/issues/exceljs/03_conditional_aggregates.json`
- Diff count: `20`
- Category: `aggregate evaluation`
- Summary: `SUMIFS`, `COUNTIFS`, and related dashboard aggregates are evaluating to `0` where Excel returns populated totals.

Representative mismatches:

- `Dashboard!B2`: `SUMIFS(Fact!$E$2:$E$11,Fact!$B$2:$B$11,$A2,Fact!$C$2:$C$11,B$1)` -> Excel `625`, werkbook `0`
- `Dashboard!C2`: `SUMIFS(Fact!$E$2:$E$11,Fact!$B$2:$B$11,$A2,Fact!$C$2:$C$11,C$1)` -> Excel `650`, werkbook `0`
- `Dashboard!B3`: `SUMIFS(Fact!$E$2:$E$11,Fact!$B$2:$B$11,$A3,Fact!$C$2:$C$11,B$1)` -> Excel `520`, werkbook `0`
- `Dashboard!C3`: `SUMIFS(Fact!$E$2:$E$11,Fact!$B$2:$B$11,$A3,Fact!$C$2:$C$11,C$1)` -> Excel `955`, werkbook `0`
