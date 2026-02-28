# Plan: WEEKNUM Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/weeknum-function-e071a2a5-bb49-4245-9caa-e8bc4cae4e5e

**Syntax:** `WEEKNUM(serial_number, [return_type])`

**Description:** Returns the week number of a specific date. The week containing January 1 is the first week of the year (System 1), or the week containing the first Thursday (ISO 8601, System 2).

**Arguments:**
- `serial_number` (required) — a date within the week
- `return_type` (optional) — determines which day the week begins on:
  - 1 or omitted: Sunday (System 1)
  - 2: Monday (System 1)
  - 11: Monday (System 1)
  - 12: Tuesday (System 1)
  - 13: Wednesday (System 1)
  - 14: Thursday (System 1)
  - 15: Friday (System 1)
  - 16: Saturday (System 1)
  - 17: Sunday (System 1)
  - 21: Monday (System 2 / ISO 8601)

**Remarks:**
- If serial_number is out of range, returns #NUM!
- If return_type is not a valid value, returns #NUM!

**Examples:**
- `=WEEKNUM(DATE(2012,3,9))` → `10` (Sunday-based)
- `=WEEKNUM(DATE(2012,3,9), 2)` → `11` (Monday-based)

## Implementation Steps

1. Add `"WEEKNUM"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnWEEKNUM(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate 1-2 args
   - Coerce arg[0] to number (serial number)
   - Default return_type to 1
   - Validate return_type is one of: 1, 2, 11-17, 21
   - Convert serial to Go `time.Time`
   - For System 1: calculate day-of-year, adjust for week start day, compute week number
   - For System 2 (return_type 21): use ISO 8601 week via `time.ISOWeek()`
3. Add `case "WEEKNUM":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `WEEKNUM(DATE(2012,3,9))` → `10` (Sunday start)
- `WEEKNUM(DATE(2012,3,9), 2)` → `11` (Monday start)
- `WEEKNUM(DATE(2012,1,1))` → `1` (Jan 1 is always week 1 in System 1)
- `WEEKNUM(DATE(2012,12,31))` → `53`
- `WEEKNUM(DATE(2012,3,9), 21)` → `10` (ISO week)
- `WEEKNUM(DATE(2012,3,9), 99)` → #NUM! (invalid return_type)
