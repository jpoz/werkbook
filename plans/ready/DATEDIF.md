# Plan: DATEDIF Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c

**Syntax:** `DATEDIF(start_date, end_date, unit)`

**Description:** Calculates the number of days, months, or years between two dates. Commonly used for age calculations.

**Arguments:**
- `start_date` (required) — the start date serial number
- `end_date` (required) — the end date serial number
- `unit` (required) — string specifying what to return:
  - `"Y"` — complete years in the period
  - `"M"` — complete months in the period
  - `"D"` — days in the period
  - `"MD"` — difference between days, ignoring months and years
  - `"YM"` — difference between months, ignoring days and years
  - `"YD"` — difference between days, ignoring years

**Remarks:**
- If start_date > end_date, returns #NUM!
- Dates are stored as serial numbers

**Examples:**
- `=DATEDIF(DATE(2001, 1, 1), DATE(2003, 1, 1), "Y")` → `2`
- `=DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "D")` → `440`
- `=DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "YD")` → `75`

## Implementation Steps

1. Add `"DATEDIF"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnDATEDIF(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 3 args
   - Coerce arg[0] and arg[1] to numbers (serial numbers)
   - Convert arg[2] to uppercase string
   - Return #NUM! if start_date > end_date
   - Convert serials to `time.Time`
   - Implement each unit:
     - **"Y"**: count complete years (compare month/day)
     - **"M"**: count complete months (years*12 + month diff, adjust for day)
     - **"D"**: simple day difference
     - **"MD"**: day difference ignoring months/years (end.Day - start.Day, handle negative by using previous month's days)
     - **"YM"**: month difference ignoring years (end.Month - start.Month, adjust for day)
     - **"YD"**: day difference ignoring years (set both to same year, compute day diff)
   - Return #NUM! for unknown unit
3. Add `case "DATEDIF":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `DATEDIF(DATE(2001, 1, 1), DATE(2003, 1, 1), "Y")` → `2`
- `DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "D")` → `440`
- `DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "YD")` → `75`
- `DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "M")` → `14`
- `DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "YM")` → `2`
- `DATEDIF(DATE(2001, 6, 1), DATE(2002, 8, 15), "MD")` → `14`
- `DATEDIF(DATE(2003, 1, 1), DATE(2001, 1, 1), "Y")` → #NUM! (start > end)
