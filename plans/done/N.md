# Plan: N Function

## Category
Info (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/n-function-a624cad1-3635-4208-b54a-29733d1278c9

**Syntax:** `N(value)`

**Description:** Converts a value to a number.

**Arguments:**
- `value` (required) — the value to convert

**Conversion rules:**
- Number → returns unchanged
- Date → returns serial number (already numeric in our system)
- TRUE → 1
- FALSE → 0
- Error → returns the error
- Everything else (text, etc.) → 0

**Examples:**
- `=N(7)` → `7`
- `=N("Even")` → `0`
- `=N(TRUE)` → `1`
- `=N("7")` → `0` (text string, not a number)

## Implementation Steps

1. Add `"N"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnN(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Switch on value type:
     - TypeNumber → return as-is
     - TypeBool → return 1 or 0
     - TypeError → return the error
     - TypeString, TypeEmpty → return 0
3. Add `case "N":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `N(7)` → `7`
- `N("text")` → `0`
- `N(TRUE)` → `1`
- `N(FALSE)` → `0`
- `N("")` → `0`
- `N(0)` → `0`
