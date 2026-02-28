# Plan: YEARFRAC Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8

**Syntax:** `YEARFRAC(start_date, end_date, [basis])`

**Description:** Calculates the fraction of the year represented by the number of whole days between two dates.

**Arguments:**
- `start_date` (required) — date serial number for start
- `end_date` (required) — date serial number for end
- `basis` (optional) — day count basis:
  - 0 or omitted: US (NASD) 30/360
  - 1: Actual/actual
  - 2: Actual/360
  - 3: Actual/365
  - 4: European 30/360

**Remarks:**
- All arguments are truncated to integers
- If start_date or end_date are not valid dates, returns #VALUE!
- If basis < 0 or basis > 4, returns #NUM!

**Examples:**
- `=YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30))` → approximately `0.58` (depends on basis)

## Implementation Steps

1. Add `"YEARFRAC"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnYEARFRAC(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate 2-3 args
   - Coerce args to numbers, truncate to integers
   - Determine basis (default 0)
   - Return #NUM! if basis < 0 or > 4
   - If start > end, swap them
   - Implement each basis:
     - **Basis 0 (US 30/360):** Adjust days per NASD rules: if d1=31, set d1=30; if d2=31 and d1>=30, set d2=30. Result = ((y2-y1)*360 + (m2-m1)*30 + (d2-d1)) / 360
     - **Basis 1 (Actual/actual):** actual days / actual days in year (handle leap years by averaging if span crosses year boundary)
     - **Basis 2 (Actual/360):** actual days / 360
     - **Basis 3 (Actual/365):** actual days / 365
     - **Basis 4 (European 30/360):** If d1=31, set d1=30; if d2=31, set d2=30. Same formula as basis 0.
3. Add `case "YEARFRAC":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30), 0)` → approximately `0.580556`
- `YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30), 1)` → approximately `0.57650`
- `YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30), 2)` → approximately `0.58611`
- `YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30), 3)` → approximately `0.57808`
- `YEARFRAC(DATE(2012, 1, 1), DATE(2012, 7, 30), 4)` → approximately `0.58056`
- `YEARFRAC(DATE(2006, 1, 1), DATE(2006, 1, 1))` → `0` (same date)
- `YEARFRAC(DATE(2006, 1, 1), DATE(2006, 1, 1), 5)` → #NUM! (invalid basis)
