# formula/UNICHAR

- Audit report: `../testdata/out/issues/formula/UNICHAR.json`
- Diff count: `4`
- Category: `unicode / control character semantics`
- Summary: `UNICHAR` disagrees with Excel on whitespace, control characters, and high code-point handling in the audited cases.

Representative mismatches:

- `Results!B1`: `_xlfn.UNICHAR(32)` -> Excel `""`, werkbook `" "`
- `Results!B10`: `_xlfn.UNICHAR(1)` -> Excel `"�"`, werkbook control char `U+0001`
- `Results!B13`: `_xlfn.UNICHAR(65535)` -> Excel `"�"`, werkbook `"￿"`
- `Results!B29`: `_xlfn.UNICHAR(TRUE)` -> Excel `"�"`, werkbook control char `U+0001`
