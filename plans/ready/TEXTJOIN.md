# Plan: TEXTJOIN Function

## Category
Text (formula/functions_text.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c

**Syntax:** `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Description:** Combines text from multiple ranges/strings with a delimiter between each value.

**Arguments:**
- `delimiter` (required) — text string or reference; numbers treated as text
- `ignore_empty` (required) — TRUE ignores empty cells, FALSE includes them
- `text1` (required) — first text item; can be string or array/range
- `text2, ...` (optional) — up to 252 additional text arguments

**Remarks:**
- If result exceeds 32767 characters, returns #VALUE!
- Empty delimiter effectively concatenates

**Examples:**
- `=TEXTJOIN(", ", TRUE, {"a","","b"})` → `"a, b"`
- `=TEXTJOIN(", ", FALSE, {"a","","b"})` → `"a, , b"`
- `=TEXTJOIN("-", TRUE, "x", "y", "z")` → `"x-y-z"`

## Implementation Steps

1. Add `"TEXTJOIN"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnTEXTJOIN(args []Value) (Value, error)` in `formula/functions_text.go`:
   - Validate at least 3 args (delimiter, ignore_empty, text1)
   - Coerce arg[0] to string for delimiter
   - Coerce arg[1] to bool for ignore_empty
   - Iterate remaining args, flattening arrays
   - If ignore_empty is true, skip empty/blank values
   - Join with delimiter, check length ≤ 32767 or return #VALUE!
3. Add `case "TEXTJOIN":` dispatch in `formula/eval.go` `callFunction`
4. Add tests in `formula/functions_text_test.go`

## Test Cases
- `TEXTJOIN(", ", TRUE, "a", "b", "c")` → `"a, b, c"`
- `TEXTJOIN(", ", TRUE, "a", "", "c")` → `"a, c"` (empty ignored)
- `TEXTJOIN(", ", FALSE, "a", "", "c")` → `"a, , c"` (empty included)
- `TEXTJOIN("", TRUE, "a", "b")` → `"ab"` (empty delimiter)
- `TEXTJOIN("-", TRUE)` → #VALUE! (too few args)
