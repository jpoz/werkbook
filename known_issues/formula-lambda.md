# formula/LAMBDA

- Audit report: `../testdata/out/issues/formula/LAMBDA.json`
- Diff count: `1`
- Category: `error typing`
- Summary: A `LAMBDA(...)` evaluation that Excel reports as an error cell is coming back as a stringified error.

Representative mismatches:

- `Results!B12`: `LAMBDA(x, x/0)(5)` -> Excel `error:#DIV/0!`, werkbook `string:#DIV/0!`
