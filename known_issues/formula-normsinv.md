# formula/NORMSINV

- Audit report: `../testdata/out/issues/formula/NORMSINV.json`
- Diff count: `6`
- Category: `numeric precision`
- Summary: `NORM.S.INV` and the round-trip checks around it are close, but still outside the exact parity expected by the audit.

Representative mismatches:

- `Results!B2`: `NORM.S.INV(0.1)` -> Excel `-1.2815515641401576`, werkbook `-1.2815515655446006`
- `Results!B3`: `NORM.S.INV(0.9)` -> Excel `1.2815515641401576`, werkbook `1.2815515655446006`
- `Results!B12`: `NORM.S.INV(Data!A2)` -> Excel `-1.2815515641401576`, werkbook `-1.2815515655446006`
- `Results!B13`: `NORM.S.INV(Data!A3)` -> Excel `1.2815515641401576`, werkbook `1.2815515655446006`
