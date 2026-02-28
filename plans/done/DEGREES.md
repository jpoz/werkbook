# Plan: DEGREES Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/degrees-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1

**Syntax:** `DEGREES(angle)`

**Description:** Converts radians to degrees.

**Arguments:**
- `angle` (required) — angle in radians to convert

**Formula:** `degrees = angle * 180 / π`

**Examples:**
- `=DEGREES(PI())` → `180`
- `=DEGREES(1)` → `57.2957795...`

## Implementation Steps

1. Add `"DEGREES"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnDEGREES(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number, return `n * 180 / math.Pi`
3. Add `case "DEGREES":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `DEGREES(PI())` → `180`
- `DEGREES(0)` → `0`
- `DEGREES(1)` → `≈57.2957795`
- `DEGREES("text")` → #VALUE!
