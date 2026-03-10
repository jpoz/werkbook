# xlsxwriter/02_lookup_model

- Audit report: `../testdata/out/issues/xlsxwriter/02_lookup_model.json`
- Diff count: `31`
- Category: `lookup / recalc`
- Summary: Lookup outputs and their dependent calculations in `Orders` are staying at `0` instead of resolving catalog values.

Representative mismatches:

- `Orders!D2`: `_xlfn.XLOOKUP(B2,Catalog!A$2:A$6,Catalog!B$2:B$6)` -> Excel `42`, werkbook `0`
- `Orders!E2`: `_xlfn.XLOOKUP(B2,Catalog!A$2:A$6,Catalog!C$2:C$6)` -> Excel `Hardware`, werkbook `0`
- `Orders!F2`: `INDEX(Catalog!D$2:D$6,MATCH(B2,Catalog!A$2:A$6,0))` -> Excel `West`, werkbook `0`
- `Orders!G2`: `C2*D2` -> Excel `126`, werkbook `0`
