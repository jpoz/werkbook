# edges/102_text_special_formats

- Audit report: `../testdata/out/issues/edges/102_text_special_formats.json`
- Diff count: `1`
- Category: `error typing`
- Summary: A special-format `TEXT(...)` case returns the textual string `#VALUE!` instead of an error value.

Representative mismatches:

- `Results!B12`: `TEXT(42,"Value: 0")` -> Excel `error:#VALUE!`, werkbook `string:#VALUE!`
