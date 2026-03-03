# Werkbook

A Go library for reading and writing Excel XLSX files with a built-in formula engine. Zero external dependencies.

## Install

### Library

```bash
go get github.com/jpoz/werkbook
```

### CLI

```bash
go install github.com/jpoz/werkbook/cmd/wb@latest
```

This installs the `wb` (werkbook) binary, which provides commands for reading, editing, and creating XLSX files from the command line:

```bash
wb info file.xlsx                        # Sheet metadata
wb read file.xlsx --range A1:D10         # Read cell data
wb edit file.xlsx --patch '[{"cell":"A1","value":"Hello"}]'  # Edit cells
wb create new.xlsx --spec '{"sheets":["Data"]}'             # Create workbook
wb calc file.xlsx                        # Recalculate formulas
wb formula list                          # List available functions
```

All output uses a JSON envelope by default. Use `--format markdown` or `--format csv` for other formats.

## Quick Start

```go
package main

import (
    "fmt"
    "log"

    "github.com/jpoz/werkbook"
)

func main() {
    // Create a new workbook
    wb := werkbook.New()
    sheet := wb.Sheet("Sheet1")

    // Set values
    sheet.SetValue("A1", "Sales")
    sheet.SetValue("A2", 100)
    sheet.SetValue("A3", 200)
    sheet.SetValue("A4", 300)

    // Set a formula
    sheet.SetFormula("A5", "SUM(A2:A4)")

    // Read the computed value
    v, _ := sheet.GetValue("A5")
    fmt.Println(v) // 600

    // Save to file
    if err := wb.SaveAs("output.xlsx"); err != nil {
        log.Fatal(err)
    }
}
```

## Reading Files

```go
wb, err := werkbook.Open("input.xlsx")
if err != nil {
    log.Fatal(err)
}

sheet := wb.Sheet("Sheet1")
v, _ := sheet.GetValue("A1")
fmt.Println(v)
```

## Formula Engine

Werkbook includes a formula engine with 184 built-in functions:

**Math (58):** ABS, ACOS, ACOSH, ARABIC, ASIN, ASINH, ATAN, ATAN2, ATANH, BASE, CEILING, COMBIN, COMBINA, COS, COSH, COT, COTH, CSC, CSCH, DECIMAL, DEGREES, EVEN, EXP, FACT, FACTDOUBLE, FLOOR, GCD, INT, LCM, LN, LOG, LOG10, MOD, MROUND, MULTINOMIAL, ODD, PERMUT, PI, POWER, PRODUCT, QUOTIENT, RADIANS, RAND, RANDBETWEEN, ROUND, ROUNDDOWN, ROUNDUP, SEC, SECH, SIGN, SIN, SINH, SQRT, SQRTPI, SUBTOTAL, TAN, TANH, TRUNC

**Statistics (30):** AVERAGE, AVEDEV, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, DEVSQ, GEOMEAN, LARGE, MAX, MAXIFS, MEDIAN, MIN, MINIFS, MODE, PERCENTILE, RANK, SMALL, STDEV, STDEVP, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, VAR, VARP

**Text (26):** CHAR, CHOOSE, CLEAN, CODE, CONCAT, CONCATENATE, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, NUMBERVALUE, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, T, TEXT, TEXTJOIN, TRIM, UPPER, VALUE

**Date & Time (22):** DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, NOW, SECOND, TIME, TODAY, WEEKDAY, WEEKNUM, WORKDAY, YEAR, YEARFRAC

**Logic (9):** AND, IF, IFERROR, IFS, NOT, OR, SORT, SWITCH, XOR

**Info (19):** COLUMN, COLUMNS, ERROR.TYPE, IFNA, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISTEXT, N, NA, ROW, ROWS, TYPE

**Lookup (8):** ADDRESS, HLOOKUP, INDEX, INDIRECT, LOOKUP, MATCH, VLOOKUP, XLOOKUP

**Finance (12):** FV, IPMT, IRR, NPER, NPV, PMT, PPMT, PV, RATE, SLN, XIRR, XNPV

## License

MIT
