# formula/RANDARRAY

- Audit report: `../testdata/out/issues/formula/RANDARRAY.json`
- Diff count: `2`
- Category: `dynamic array spill`
- Summary: Spill results expected from `RANDARRAY` are missing from downstream cells in the audited workbook.

Representative mismatches:

- `Results!C3`: spilled value expected `0.6795156853488269`, werkbook `empty:null`
- `Results!D3`: spilled value expected `0.965946570656415`, werkbook `empty:null`
