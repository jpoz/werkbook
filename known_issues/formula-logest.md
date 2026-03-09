# formula/LOGEST

- Audit report: `../testdata/out/issues/formula/LOGEST.json`
- Diff count: `3`
- Category: `statistical function`
- Summary: `LOGEST` has both numeric divergence and one case where werkbook returns a number while Excel reports `#NUM!`.

Representative mismatches:

- `Results!B11`: `INDEX(LOGEST(Data!B1:B5,Data!A1:A5,,TRUE),4,1)` -> Excel `2.7514573619570592E+32`, werkbook `5.568425613484525E+31`
- `Results!B30`: `INDEX(LOGEST(Data!B1:B5,Data!A1:A5,FALSE,TRUE),4,1)` -> Excel `error:#NUM!`, werkbook `4.2876877223830837E+33`
- `Results!B44`: `INDEX(LOGEST(Data!C1:C5,Data!A1:A5,,TRUE),4,1)` -> Excel `9.699199525497484E+31`, werkbook `2.462687379520846E+31`
