# Plan: NETWORKDAYS Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7

**Syntax:** `NETWORKDAYS(start_date, end_date, [holidays])`

**Description:** Returns the number of whole working days between start_date and end_date. Working days exclude weekends (Saturday and Sunday) and any dates identified in holidays.

**Arguments:**
- `start_date` (required) — date serial number for start
- `end_date` (required) — date serial number for end
- `holidays` (optional) — range/array of date serial numbers to exclude as holidays

**Remarks:**
- If any argument is not a valid date, returns #VALUE!
- Both start_date and end_date are inclusive
- If start_date > end_date, result is negative

**Examples:**
- `=NETWORKDAYS(DATE(2012,10,1), DATE(2013,3,1))` → `110`
- With 1 holiday excluded → `109`
- With 3 holidays excluded → `107`

## Implementation Steps

1. Add `"NETWORKDAYS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnNETWORKDAYS(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate 2-3 args
   - Coerce arg[0] and arg[1] to numbers (serial numbers)
   - If arg[2] provided, flatten array to get holiday serial numbers
   - Convert start/end to Go `time.Time`
   - Build a set of holiday serial numbers for fast lookup
   - Iterate from start to end (or end to start if reversed), counting days where:
     - `time.Weekday()` is not Saturday or Sunday
     - Serial number is not in the holiday set
   - Handle direction: if start > end, count is negative
3. Add `case "NETWORKDAYS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `NETWORKDAYS(DATE(2012,10,1), DATE(2013,3,1))` → `110`
- `NETWORKDAYS(DATE(2012,10,1), DATE(2012,10,1))` → `1` (same day, weekday)
- `NETWORKDAYS(DATE(2012,10,6), DATE(2012,10,6))` → `0` (Saturday)
- `NETWORKDAYS(DATE(2013,3,1), DATE(2012,10,1))` → `-110` (reversed)
- `NETWORKDAYS(DATE(2012,10,1), DATE(2012,10,5))` → `5` (Mon-Fri)
