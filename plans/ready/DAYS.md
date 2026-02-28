# Plan: DAYS Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/days-function-57740535-d549-4395-8728-0f07bff0b9df

**Syntax:** `DAYS(end_date, start_date)`

**Description:** Returns the number of days between two dates.

**Arguments:**
- `end_date` (required) — the later date
- `start_date` (required) — the earlier date

**Remarks:**
- When both are numbers, simply returns end_date - start_date
- Result can be negative if start_date > end_date
- Dates are Excel serial numbers

**Examples:**
- `=DAYS(DATE(2021,3,15), DATE(2021,2,1))` → `42`
- `=DAYS(DATE(2021,12,31), DATE(2021,1,1))` → `364`

## Implementation Steps

1. Add `"DAYS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnDAYS(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 2 args
   - Coerce both to numbers (serial numbers)
   - Return `math.Trunc(end) - math.Trunc(start)`
3. Add `case "DAYS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `DAYS(DATE(2021,3,15), DATE(2021,2,1))` → `42`
- `DAYS(DATE(2021,12,31), DATE(2021,1,1))` → `364`
- `DAYS(DATE(2020,1,1), DATE(2021,1,1))` → `-366` (leap year)
- `DAYS(100, 50)` → `50`
