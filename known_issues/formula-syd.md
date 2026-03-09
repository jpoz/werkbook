# formula/SYD

- Audit report: `../testdata/out/issues/formula/SYD.json`
- Diff count: `3`
- Category: `argument validation / return handling`
- Summary: `SYD` diverges in both invalid-argument handling and normal result production.

Representative mismatches:

- `Results!B42`: `IFERROR(SYD(10000,1000,5),"err")` -> Excel `string:err`, werkbook `8000`
- `Results!B43`: `IFERROR(SYD(10000,1000,5,1,1),"err")` -> Excel `string:err`, werkbook `empty:null`
- `Results!B44`: `SYD(10000,2000,5,1)+SYD(10000,2000,5,2)+SYD(10000,2000,5,3)+SYD(10000,2000,5,4)+SYD(10000,2000,5,5)` -> Excel `8000`, werkbook `empty:null`
