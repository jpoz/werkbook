# formula/LET

- Audit report: `../testdata/out/issues/formula/LET.json`
- Diff count: `37`
- Category: `feature / baseline mismatch`
- Summary: The checked-in spec expects `#NAME?` for `LET(...)`, while werkbook is evaluating the formulas. This looks more like a parity-baseline disagreement than a missing implementation.

Representative mismatches:

- `Results!B1`: `LET(x,5,x)` -> Excel `error:#NAME?`, werkbook `5`
- `Results!B2`: `LET(x,5,x+1)` -> Excel `error:#NAME?`, werkbook `6`
- `Results!B3`: `LET(x,5,y,2,x+y)` -> Excel `error:#NAME?`, werkbook `7`
- `Results!B4`: `LET(x,5,y,x*2,y)` -> Excel `error:#NAME?`, werkbook `10`
