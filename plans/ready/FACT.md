# Plan: FACT Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3

**Syntax:** `FACT(number)`

**Description:** Returns the factorial of a number (1×2×3×...×number).

**Arguments:**
- `number` (required) — nonnegative number; non-integers are truncated

**Remarks:**
- Negative numbers return #NUM!
- FACT(0) = 1
- Non-integer values truncated before computing

**Examples:**
- `=FACT(5)` → `120`
- `=FACT(1.9)` → `1` (truncated to FACT(1))
- `=FACT(0)` → `1`
- `=FACT(-1)` → #NUM!

## Implementation Steps

1. Add `"FACT"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnFACT(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number, truncate to integer
   - If negative, return #NUM!
   - Compute factorial iteratively (or use `math.Gamma(n+1)` for large values)
3. Add `case "FACT":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `FACT(5)` → `120`
- `FACT(0)` → `1`
- `FACT(1)` → `1`
- `FACT(1.9)` → `1`
- `FACT(10)` → `3628800`
- `FACT(-1)` → #NUM!
- `FACT("text")` → #VALUE!
