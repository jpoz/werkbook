# Plan: GCD Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/gcd-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a

**Syntax:** `GCD(number1, [number2], ...)`

**Description:** Returns the greatest common divisor of two or more integers.

**Arguments:**
- `number1` (required) — first value
- `number2, ...` (optional) — up to 255 values; non-integers truncated

**Remarks:**
- Non-numeric args return #VALUE!
- Negative args return #NUM!
- Args ≥ 2^53 return #NUM!
- GCD(n, 0) = n
- GCD(0, 0) = 0

**Examples:**
- `=GCD(5, 2)` → `1`
- `=GCD(24, 36)` → `12`
- `=GCD(7, 1)` → `1`
- `=GCD(5, 0)` → `5`

## Implementation Steps

1. Add `"GCD"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnGCD(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate at least 1 arg
   - Coerce each arg to number, truncate to integer
   - Return #NUM! if any arg is negative
   - Use Euclidean algorithm iteratively across all args
3. Add `case "GCD":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `GCD(5, 2)` → `1`
- `GCD(24, 36)` → `12`
- `GCD(7, 1)` → `1`
- `GCD(5, 0)` → `5`
- `GCD(0, 0)` → `0`
- `GCD(12, 8, 4)` → `4`
- `GCD(-5, 2)` → #NUM!
