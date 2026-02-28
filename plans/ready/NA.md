# Plan: NA Function

## Category
Info (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/na-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c

**Syntax:** `NA()`

**Description:** Returns the #N/A error value. Used to mark cells as intentionally empty.

**Arguments:**
- None. Takes no parameters.

**Remarks:**
- Empty parentheses required
- Equivalent to typing #N/A directly in a cell
- Formulas referencing a cell containing #N/A will also return #N/A

## Implementation Steps

1. Add `"NA"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnNA(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 0 args (return #VALUE! if args provided)
   - Return `ErrorVal(ErrValNA)`
3. Add `case "NA":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `NA()` → #N/A
