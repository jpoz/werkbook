# Plan: DATEVALUE Function

## Category
Date/Time (formula/functions_date.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252

**Syntax:** `DATEVALUE(date_text)`

**Description:** Converts a date stored as text to a serial number that Excel recognizes as a date.

**Arguments:**
- `date_text` (required) — text representing a date in a recognized format

**Remarks:**
- Must represent a date between January 1, 1900 and December 31, 9999
- Returns #VALUE! if text cannot be parsed as a date
- If year is omitted, uses current year
- Time information in date_text is ignored

**Supported formats (common):**
- `"1/30/2008"` (M/D/YYYY)
- `"30-Jan-2008"` (D-Mon-YYYY)
- `"2008/01/30"` (YYYY/MM/DD)
- `"January 30, 2008"` (Month D, YYYY)

**Examples:**
- `=DATEVALUE("8/22/2011")` → `40777`
- `=DATEVALUE("22-MAY-2011")` → `40685`
- `=DATEVALUE("2011/02/23")` → `40597`

## Implementation Steps

1. Add `"DATEVALUE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnDATEVALUE(args []Value) (Value, error)` in `formula/functions_date.go`:
   - Validate exactly 1 arg
   - Coerce arg[0] to string
   - Try parsing with multiple Go time layouts:
     - `"1/2/2006"`, `"01/02/2006"` (M/D/YYYY)
     - `"2-Jan-2006"`, `"02-Jan-2006"` (D-Mon-YYYY)
     - `"2006/01/02"`, `"2006-01-02"` (YYYY/MM/DD, YYYY-MM-DD)
     - `"January 2, 2006"` (Month D, YYYY)
   - If no format matches, return #VALUE!
   - Convert parsed time to Excel serial using `timeToExcelSerial()`
   - Return as number Value (floor — strip time component)
3. Add `case "DATEVALUE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_date_test.go`

## Test Cases
- `DATEVALUE("8/22/2011")` → `40777`
- `DATEVALUE("22-MAY-2011")` → `40685`
- `DATEVALUE("2011/02/23")` → `40597`
- `DATEVALUE("January 1, 2008")` → `39448`
- `DATEVALUE("not a date")` → #VALUE!
