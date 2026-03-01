# Werkbook

A Go library for reading and writing Excel XLSX files with a built-in formula engine. Zero external dependencies.

## Install

### Library

```bash
go get github.com/werkbook/werkbook
```

### CLI

```bash
go install github.com/werkbook/werkbook/cmd/werkbook@latest
```

This installs the `werkbook` binary, which provides commands for reading, editing, and creating XLSX files from the command line:

```bash
werkbook info file.xlsx                        # Sheet metadata
werkbook read file.xlsx --range A1:D10         # Read cell data
werkbook edit file.xlsx --patch '[{"cell":"A1","value":"Hello"}]'  # Edit cells
werkbook create new.xlsx --spec '{"sheets":["Data"]}'             # Create workbook
werkbook calc file.xlsx                        # Recalculate formulas
werkbook formula list                          # List available functions
```

All output uses a JSON envelope by default. Use `--format markdown` or `--format csv` for other formats.

## Quick Start

```go
package main

import (
    "fmt"
    "log"

    "github.com/werkbook/werkbook"
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

Werkbook includes a formula engine with 55 built-in functions:

**Math (12):** ABS, CEILING, FLOOR, INT, MOD, POWER, RAND, RANDBETWEEN, ROUND, ROUNDDOWN, ROUNDUP, SQRT

**Statistics (15):** AVERAGE, AVERAGEIF, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, LARGE, MAX, MIN, SMALL, SUM, SUMIF, SUMIFS, SUMPRODUCT

**Text (13):** CHOOSE, CONCATENATE, CONCAT, FIND, LEFT, LEN, LOWER, MID, RIGHT, SUBSTITUTE, TEXT, TRIM, UPPER

**Date (6):** DATE, DAY, MONTH, NOW, TODAY, YEAR

**Logic (5):** AND, IF, IFERROR, NOT, OR

**Lookup (4):** HLOOKUP, INDEX, MATCH, VLOOKUP

**Info (1):** IFNA

## License

MIT
