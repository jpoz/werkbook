# Plan: WORKDAY Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33

**Syntax:** `WORKDAY(start_date, days, [holidays])`

**Description:** Returns the serial number of the date that is the indicated number of working days before or after a start date. Working days exclude weekends (Saturday/Sunday) and any specified holidays. This is the inverse of NETWORKDAYS.

**Arguments:**
- `start_date` (required) — the starting date serial number
- `days` (required) — number of nonweekend/nonholiday days before or after start_date; positive = future, negative = past; truncated to integer
- `holidays` (optional) — range/array of date serial numbers to exclude as holidays

**Remarks:**
- If any argument is not a valid date, returns #VALUE!
- If start_date plus days yields an invalid date, returns #NUM!
- If days is not an integer, it is truncated

**Examples:**
- `=WORKDAY(DATE(2008, 10, 1), 151)` → serial number for `2009-04-30`

## Implementation Steps

1. Add `"WORKDAY"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnWORKDAY(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate 2-3 args
   - Coerce arg[0] to number (start_date serial), arg[1] to number (days, truncate to int)
   - If arg[2] provided, flatten array to get holiday serial numbers into a set
   - Convert start serial to `time.Time`
   - Step forward (or backward if days < 0) one day at a time:
     - Skip weekends (Saturday/Sunday)
     - Skip holidays
     - Decrement remaining day count for each valid workday
   - Convert resulting `time.Time` back to Excel serial number
   - Return as NumberVal
3. Add `case "WORKDAY":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `WORKDAY(DATE(2008, 10, 1), 151)` → `DATE(2009, 4, 30)` (serial 39933)
- `WORKDAY(DATE(2008, 10, 1), 0)` → `DATE(2008, 10, 1)` (same day, Wednesday)
- `WORKDAY(DATE(2008, 10, 3), 1)` → `DATE(2008, 10, 6)` (Friday + 1 = Monday)
- `WORKDAY(DATE(2008, 10, 6), -1)` → `DATE(2008, 10, 3)` (Monday - 1 = Friday)
- `WORKDAY(DATE(2008, 10, 1), 5)` → `DATE(2008, 10, 8)` (Wed + 5 workdays = Wed)
