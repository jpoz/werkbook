# Plan: ISLOGICAL Function

## Category
Information (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665

**Syntax:** `ISLOGICAL(value)`

**Description:** Returns TRUE if value refers to a logical value (TRUE or FALSE).

**Arguments:**
- `value` (required) — the value to test

**Remarks:**
- Value arguments are not converted — `"TRUE"` as text returns FALSE
- Returns TRUE only for actual boolean TRUE/FALSE values
- Returns FALSE for numbers, text, errors, blanks

**Examples:**
- `=ISLOGICAL(TRUE)` → `TRUE`
- `=ISLOGICAL("TRUE")` → `FALSE`
- `=ISLOGICAL(1)` → `FALSE`
- `=ISLOGICAL(1>0)` → `TRUE`

## Implementation Steps

1. Add `"ISLOGICAL"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnISLOGICAL(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Check if arg[0].Type == TypeBool
   - Return BoolValue(true) if so, BoolValue(false) otherwise
3. Add `case "ISLOGICAL":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `ISLOGICAL(TRUE)` → `TRUE`
- `ISLOGICAL(FALSE)` → `TRUE`
- `ISLOGICAL("TRUE")` → `FALSE`
- `ISLOGICAL(1)` → `FALSE`
- `ISLOGICAL(0)` → `FALSE`
- `ISLOGICAL("")` → `FALSE`
- `ISLOGICAL(1/0)` → `FALSE` (error value)
