# Plan: EDATE Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5

**Syntax:** `EDATE(start_date, months)`

**Description:** Returns the serial number that represents the date that is the indicated number of months before or after a specified date. Use EDATE to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.

**Arguments:**
- `start_date` (required) — a date serial number representing the start date
- `months` (required) — number of months before or after start_date; positive = future, negative = past

**Remarks:**
- If start_date is not a valid date, returns #VALUE!
- If months is not an integer, it is truncated
- Uses Excel serial number date system (Jan 1, 1900 = 1)

**Examples:**
- `=EDATE(DATE(2011,1,15), 1)` → serial for 15-Feb-2011
- `=EDATE(DATE(2011,1,15), -1)` → serial for 15-Dec-2010
- `=EDATE(DATE(2011,1,15), 2)` → serial for 15-Mar-2011

## Implementation Steps

1. Add `"EDATE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnEDATE(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 2 args
   - Coerce arg[0] to number (serial number)
   - Coerce arg[1] to number, truncate to integer for months
   - Convert serial to Go `time.Time` using `excelSerialToTime()`
   - Use `time.AddDate(0, months, 0)` to shift by months
   - If resulting day overflows month (e.g. Jan 31 + 1 month), clamp to last day of target month
   - Convert back to serial with `timeToExcelSerial()`
3. Add `case "EDATE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `EDATE(DATE(2011,1,15), 1)` → serial for 2011-02-15
- `EDATE(DATE(2011,1,15), -1)` → serial for 2010-12-15
- `EDATE(DATE(2011,1,15), 2)` → serial for 2011-03-15
- `EDATE(DATE(2011,1,31), 1)` → serial for 2011-02-28 (clamped)
- `EDATE(DATE(2020,1,29), 1)` → serial for 2020-02-29 (leap year)
- `EDATE(DATE(2011,1,15), 0)` → serial for 2011-01-15 (no change)
