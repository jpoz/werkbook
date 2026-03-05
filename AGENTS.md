# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Build & Test Commands

```bash
# Run all unit tests
gotestsum ./...

# Run a single test
gotestsum -- -run TestRoundTrip ./...
```

No external dependencies beyond the Go standard library. Module: `github.com/jpoz/werkbook`.

## Useful Tools

```bash
# Look up Excel function documentation (e.g. syntax, behavior, edge cases)
exceldoc SUM
```

## Architecture

Werkbook is a Go library for reading and writing Excel XLSX files with a built-in formula engine supporting ~170 functions.

### Two-layer design

- **Public API layer** (root package `werkbook`): `File`, `Sheet`, `Row`, `Cell`, `Value` types. All coordinates are 1-based. Cells are stored sparsely in `map[int]*Row` -> `map[int]*Cell`.
- **OOXML layer** (`ooxml/` package): Handles ZIP/XML serialization. Communicates with the public layer through the `WorkbookData`/`SheetData`/`RowData`/`CellData` intermediary types defined in `ooxml/workbook.go`. The public API never exposes OOXML internals.

### Data flow

- **Write path**: `File.buildWorkbookData()` converts `Sheet` -> `SheetData`, then `ooxml.WriteWorkbook()` serializes to ZIP. Strings use a shared string table (SST); cells with `Type: "s"` have their raw string replaced with an SST index during write.
- **Read path**: `ooxml.ReadWorkbook()` parses ZIP -> `WorkbookData`, then `fileFromData()` converts back to `File`/`Sheet`/`Cell` structures. SST indices are resolved to actual strings during read.

### Value system

`Value` is a tagged union (`value.go`) with types: `TypeEmpty`, `TypeNumber`, `TypeString`, `TypeBool`, `TypeError`. The `toValue()` function converts Go types (all int/uint/float variants, string, bool, `time.Time`, nil) to `Value`. Dates are stored as Excel serial numbers via `timeToExcelSerial()` in `date.go`, which accounts for the Excel 1900 leap year bug.

### Coordinate system

`coords.go` handles A1-style references. `CellNameToCoordinates("B3")` returns `(col=2, row=3)`. Max: 1,048,576 rows, 16,384 columns (XFD).
