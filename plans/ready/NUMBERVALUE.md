# Plan: NUMBERVALUE Function

## Category
Text (formula/functions_text.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/numbervalue-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879

**Syntax:** `NUMBERVALUE(text, [decimal_separator], [group_separator])`

**Description:** Converts text to a number, in a locale-independent way.

**Arguments:**
- `text` (required) — the text to convert to a number
- `decimal_separator` (optional) — character used to separate integer and fractional parts; default is "."
- `group_separator` (optional) — character used to separate groupings (e.g., thousands); default is ","

**Remarks:**
- If multiple characters are used for separators, only the first character is used
- If text is empty string (""), result is 0
- Spaces in text are ignored, even in the middle
- If decimal separator is used more than once, returns #VALUE!
- If group separator occurs after the decimal separator, returns #VALUE!
- Leading/trailing percent signs are handled: "50%" → 0.5

**Examples:**
- `=NUMBERVALUE("2.500,27", ",", ".")` → `2500.27`
- `=NUMBERVALUE("3.5%")` → `0.035`
- `=NUMBERVALUE("")` → `0`

## Implementation Steps

1. Add `"NUMBERVALUE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnNUMBERVALUE(args []Value) (Value, error)` in `formula/functions_text.go`:
   - Validate 1-3 args
   - Convert arg[0] to string
   - Get decimal separator from arg[1] (default "."), take first char
   - Get group separator from arg[2] (default ","), take first char
   - Strip all spaces from text
   - Count and handle percent signs (each "%" divides result by 100)
   - Remove all group separator characters (but return #VALUE! if group separator appears after decimal separator)
   - Validate only one decimal separator exists
   - Replace decimal separator with "." for Go parsing
   - Parse with `strconv.ParseFloat`
   - Apply percent scaling
   - Return as NumberVal
3. Add `case "NUMBERVALUE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_text_test.go`
5. Add `"NUMBERVALUE": "_xlfn."` to `xlfnPrefix` map in `formula/xlfn.go`

## Test Cases
- `NUMBERVALUE("2.500,27", ",", ".")` → `2500.27`
- `NUMBERVALUE("3.5%")` → `0.035`
- `NUMBERVALUE("")` → `0`
- `NUMBERVALUE("  3 000  ")` → `3000`
- `NUMBERVALUE("1,234.56")` → `1234.56`
- `NUMBERVALUE("1.234,56", ",", ".")` → `1234.56`
- `NUMBERVALUE("1.2.3")` → #VALUE! (multiple decimal separators)
- `NUMBERVALUE("50%%")` → `0.005` (two percent signs)
