# Plan: WEEKDAY Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a

**Syntax:** `WEEKDAY(serial_number, [return_type])`

**Description:** Returns the day of the week as an integer.

**Arguments:**
- `serial_number` (required) — Excel date serial number
- `return_type` (optional) — numbering system:
  - 1 or omitted: 1 (Sunday) through 7 (Saturday)
  - 2: 1 (Monday) through 7 (Sunday)
  - 3: 0 (Monday) through 6 (Sunday)
  - 11: 1 (Monday) through 7 (Sunday)
  - 12: 1 (Tuesday) through 7 (Monday)
  - 13: 1 (Wednesday) through 7 (Tuesday)
  - 14: 1 (Thursday) through 7 (Wednesday)
  - 15: 1 (Friday) through 7 (Thursday)
  - 16: 1 (Saturday) through 7 (Friday)
  - 17: 1 (Sunday) through 7 (Saturday)

**Remarks:**
- Invalid serial_number or return_type returns #NUM!
- Use existing `excelSerialToTime()` to convert serial → Go time.Time, then get weekday

**Examples:**
- `=WEEKDAY(DATE(2008,2,14))` → `5` (Thursday, Sunday-based)
- `=WEEKDAY(DATE(2008,2,14), 2)` → `4` (Thursday, Monday-based)
- `=WEEKDAY(DATE(2008,2,14), 3)` → `3` (Thursday, 0-based Monday)

## Implementation Steps

1. Add `"WEEKDAY"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnWEEKDAY(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate 1-2 args
   - Coerce arg[0] to number (serial number)
   - Default return_type to 1
   - Convert serial to Go time.Time using existing helper
   - Get Go's time.Weekday() (0=Sunday...6=Saturday)
   - Map to requested return_type numbering
3. Add `case "WEEKDAY":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `WEEKDAY(DATE(2008,2,14))` → `5` (Thursday, type 1)
- `WEEKDAY(DATE(2008,2,14), 2)` → `4` (Thursday, type 2)
- `WEEKDAY(DATE(2008,2,14), 3)` → `3` (Thursday, type 3)
- `WEEKDAY(DATE(2024,1,1))` → `2` (Monday, type 1)
- `WEEKDAY(DATE(2024,1,1), 2)` → `1` (Monday, type 2)
