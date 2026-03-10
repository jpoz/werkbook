# xlsxwriter/10_future_function_prefix_shape

- Audit report: `../testdata/out/issues/xlsxwriter/10_future_function_prefix_shape.json`
- Diff count: `6`
- Category: `future-function prefix semantics`
- Summary: The checked-in spec reports `#NAME?` for `_xlfn` and `_xludf` prefixed formulas, while werkbook strips the prefixes and evaluates the formulas successfully.

Representative mismatches:

- `Orders!F2`: `_xlfn.SWITCH(_xludf.IFNA(_xlfn.XLOOKUP(B2,Catalog!A$2:A$6,Catalog!D$2:D$6),"Missing"),"West","Mountain","East","Atlantic","Central","Central","Missing","Review","General")` -> Excel `error:#NAME?`, werkbook `General`
- `Orders!F3`: same formula pattern -> Excel `error:#NAME?`, werkbook `General`
- `Orders!F4`: same formula pattern -> Excel `error:#NAME?`, werkbook `General`
- `Orders!F5`: same formula pattern -> Excel `error:#NAME?`, werkbook `General`
