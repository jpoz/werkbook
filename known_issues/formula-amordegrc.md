# formula/AMORDEGRC

- Audit report: `../testdata/out/issues/formula/AMORDEGRC.json`
- Diff count: `1`
- Category: `rounding / off by one`
- Summary: `AMORDEGRC` is off by one on at least one audited depreciation case.

Representative mismatches:

- `Results!B29`: `AMORDEGRC(Data!C1,Data!C2,Data!C3,Data!C4,3,Data!C5,1)` -> Excel `5273`, werkbook `5274`
