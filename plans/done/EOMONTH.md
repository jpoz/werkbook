# Plan: EOMONTH Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628

**Syntax:** `EOMONTH(start_date, months)`

**Description:** Returns the serial number for the last day of the month that is the indicated number of months before or after start_date. Use EOMONTH to calculate maturity dates or due dates that fall on the last day of the month.

**Arguments:**
- `start_date` (required) — a date serial number representing the starting date
- `months` (required) — number of months before or after start_date; positive = future, negative = past

**Remarks:**
- If start_date is not a valid date, returns #NUM!
- If start_date plus months yields an invalid date, returns #NUM!
- If months is not an integer, it is truncated

**Examples:**
- `=EOMONTH(DATE(2011,1,1), 1)` → serial for 2011-02-28
- `=EOMONTH(DATE(2011,1,1), -3)` → serial for 2010-10-31

## Implementation Steps

1. Add `"EOMONTH"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnEOMONTH(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 2 args
   - Coerce arg[0] to number (serial number)
   - Coerce arg[1] to number, truncate to integer for months
   - Convert serial to Go `time.Time` using `excelSerialToTime()`
   - Shift by months, then find last day of that month:
     - Go to first day of target month, then `AddDate(0, 1, -1)`
   - Convert back to serial with `timeToExcelSerial()`
3. Add `case "EOMONTH":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `EOMONTH(DATE(2011,1,1), 1)` → serial for 2011-02-28
- `EOMONTH(DATE(2011,1,1), -3)` → serial for 2010-10-31
- `EOMONTH(DATE(2011,1,1), 0)` → serial for 2011-01-31
- `EOMONTH(DATE(2020,1,1), 1)` → serial for 2020-02-29 (leap year)
- `EOMONTH(DATE(2011,12,1), 1)` → serial for 2012-01-31 (year rollover)
