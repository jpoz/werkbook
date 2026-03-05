# Werkbook

[![Go](https://github.com/jpoz/werkbook/actions/workflows/go.yml/badge.svg)](https://github.com/jpoz/werkbook/actions/workflows/go.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/jpoz/werkbook.svg)](https://pkg.go.dev/github.com/jpoz/werkbook)
[![Go Report Card](https://goreportcard.com/badge/github.com/jpoz/werkbook)](https://goreportcard.com/report/github.com/jpoz/werkbook)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

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

Werkbook includes a formula engine with 198 built-in functions. See [FORMULAS.md](FORMULAS.md) for the full list of supported and unsupported functions.

## License

MIT
