# edges/59_type_n_t_coercion

- Audit report: `../testdata/out/issues/edges/59_type_n_t_coercion.json`
- Diff count: `1`
- Category: `error typing`
- Summary: `T(...)` is returning an error-looking string where Excel keeps the underlying error as an error cell.

Representative mismatches:

- `Results!B27`: `T(Data!A6)` -> Excel `error:#DIV/0!`, werkbook `string:#DIV/0!`
