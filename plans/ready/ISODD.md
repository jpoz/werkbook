# Plan: ISODD Function

## Category
Info (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665

**Syntax:** `ISODD(number)`

**Description:** Returns TRUE if number is odd, FALSE if even.

**Arguments:**
- `number` (required) — value to test; non-integers truncated

**Remarks:**
- Non-numeric values return #VALUE!
- Complement of ISEVEN

**Examples:**
- `=ISODD(1)` → `TRUE`
- `=ISODD(2)` → `FALSE`
- `=ISODD(2.5)` → `FALSE` (truncated to 2)
- `=ISODD(0)` → `FALSE`

## Implementation Steps

1. Add `"ISODD"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnISODD(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number; return #VALUE! if not numeric
   - Truncate to integer, check `int(n) % 2 != 0`
   - Return BoolVal
3. Add `case "ISODD":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `ISODD(1)` → `TRUE`
- `ISODD(2)` → `FALSE`
- `ISODD(0)` → `FALSE`
- `ISODD(2.5)` → `FALSE`
- `ISODD(-3)` → `TRUE`
- `ISODD("text")` → #VALUE!
