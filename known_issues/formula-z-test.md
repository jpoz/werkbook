# formula/Z_TEST

- Audit report: `../testdata/out/issues/formula/Z_TEST.json`
- Diff count: `1`
- Category: `numeric precision`
- Summary: `Z.TEST` is very close but still not bit-for-bit aligned with the Excel-backed spec.

Representative mismatches:

- `Results!B13`: `_xlfn.Z.TEST(Data!A1:A10,0)` -> Excel `2.825439313895175E-10`, werkbook `2.8254398820592996E-10`
