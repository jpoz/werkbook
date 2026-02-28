# Plan: ISEVEN Function

## Category
Info (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356

**Syntax:** `ISEVEN(number)`

**Description:** Returns TRUE if number is even, FALSE if odd.

**Arguments:**
- `number` (required) — value to test; non-integers truncated

**Remarks:**
- Non-numeric values return #VALUE!
- Zero is considered even (ISEVEN(0) → TRUE)

**Examples:**
- `=ISEVEN(-1)` → `FALSE`
- `=ISEVEN(2.5)` → `TRUE` (truncated to 2)
- `=ISEVEN(5)` → `FALSE`
- `=ISEVEN(0)` → `TRUE`

## Implementation Steps

1. Add `"ISEVEN"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnISEVEN(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number; return #VALUE! if not numeric
   - Truncate to integer, check `int(n) % 2 == 0`
   - Return BoolVal
3. Add `case "ISEVEN":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `ISEVEN(0)` → `TRUE`
- `ISEVEN(1)` → `FALSE`
- `ISEVEN(2)` → `TRUE`
- `ISEVEN(2.5)` → `TRUE`
- `ISEVEN(-1)` → `FALSE`
- `ISEVEN(-4)` → `TRUE`
- `ISEVEN("text")` → #VALUE!
