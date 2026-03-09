# formula/NORMSDIST

- Audit report: `../testdata/out/issues/formula/NORMSDIST.json`
- Diff count: `1`
- Category: `future function alias / error handling`
- Summary: The audited `NORM.S.DIST` case falls through to blank in werkbook, so `IFERROR(...)` does not produce the expected `"err"` result.

Representative mismatches:

- `Results!B47`: `IFERROR(NORM.S.DIST(1),"err")` -> Excel `string:err`, werkbook `empty:null`
