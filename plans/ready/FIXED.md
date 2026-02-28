# Plan: FIXED Function

## Category
Text (formula/functions_text.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/fixed-function-ffd5723c-324c-45e9-8b96-e41be2a8274a

**Syntax:** `FIXED(number, [decimals], [no_commas])`

**Description:** Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.

**Arguments:**
- `number` (required) — the number to round and convert to text
- `decimals` (optional) — number of digits to the right of the decimal point; default 2
- `no_commas` (optional) — if TRUE, prevents commas in the result; default FALSE

**Remarks:**
- If decimals is negative, number is rounded to the left of the decimal point
- Decimals can be as large as 127
- Numbers can never have more than 15 significant digits

**Examples:**
- `=FIXED(1234.567, 1)` → `"1,234.6"`
- `=FIXED(1234.567, -1)` → `"1,230"`
- `=FIXED(-1234.567, -1, TRUE)` → `"-1230"`
- `=FIXED(44.332)` → `"44.33"`

## Implementation Steps

1. Add `"FIXED"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnFIXED(args []Value) (Value, error)` in `formula/functions_text.go`:
   - Validate 1-3 args
   - Coerce arg[0] to number
   - Default decimals to 2, no_commas to false
   - Round number to specified decimals (handle negative decimals by rounding to left of decimal)
   - Format with `fmt.Sprintf` for the decimal portion
   - If no_commas is false, insert commas every 3 digits in the integer part
   - Return as string Value
3. Add `case "FIXED":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_text_test.go`

## Test Cases
- `FIXED(1234.567, 1)` → `"1,234.6"`
- `FIXED(1234.567, -1)` → `"1,230"`
- `FIXED(-1234.567, -1, TRUE)` → `"-1230"`
- `FIXED(44.332)` → `"44.33"` (default 2 decimals)
- `FIXED(0.5, 0)` → `"1"` (rounds up)
- `FIXED(1234567.89, 2)` → `"1,234,567.89"`
- `FIXED(1234567.89, 2, TRUE)` → `"1234567.89"` (no commas)
