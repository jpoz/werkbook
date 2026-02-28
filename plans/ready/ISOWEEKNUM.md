# Plan: ISOWEEKNUM Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e

**Syntax:** `ISOWEEKNUM(date)`

**Description:** Returns the ISO week number of the year for a given date. ISO 8601 defines week 1 as the week containing the first Thursday of the year.

**Arguments:**
- `date` (required) — date serial number

**Remarks:**
- If the date argument is not a valid number, returns #NUM!
- If the date argument is not a valid date type, returns #VALUE!
- Go's `time.ISOWeek()` directly provides this functionality

**Examples:**
- `=ISOWEEKNUM(DATE(2012, 3, 9))` → `10`

## Implementation Steps

1. Add `"ISOWEEKNUM"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnISOWEEKNUM(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 1 arg
   - Coerce to number (serial number)
   - Convert serial number to `time.Time` using `excelSerialToTime()`
   - Use `t.ISOWeek()` to get the ISO week number
   - Return the week number as NumberVal
3. Add `case "ISOWEEKNUM":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`
5. Add `"ISOWEEKNUM": "_xlfn."` to `xlfnPrefix` map in `formula/xlfn.go`

## Test Cases
- `ISOWEEKNUM(DATE(2012, 3, 9))` → `10`
- `ISOWEEKNUM(DATE(2023, 1, 1))` → `52` (Jan 1 2023 is a Sunday, belongs to week 52 of 2022 by ISO)
- `ISOWEEKNUM(DATE(2023, 1, 2))` → `1` (first Monday of 2023)
- `ISOWEEKNUM(DATE(2020, 12, 31))` → `53` (2020 has 53 ISO weeks)
