# wb CLI

A command-line tool for reading, writing, and manipulating Excel `.xlsx` files.

## Install

```bash
go install github.com/jpoz/werkbook/cmd/wb@latest
```

## Usage

```
wb <command> [flags] <file>
```

### Global Flags

| Flag | Description |
|------|-------------|
| `--format <json\|markdown\|csv>` | Output format (default: `json`) |
| `--mode <default\|agent>` | Output contract mode (default: `default`) |
| `--compact` | Emit compact JSON (no indentation) |

## Commands

### `capabilities` — Show machine-readable CLI metadata

```bash
wb capabilities
```

Returns structured metadata for commands, flags, modes, and agent-mode behavior.
This is the preferred discovery entrypoint for agentic usage.

### `info` — Show workbook metadata

```bash
wb info data.xlsx
wb info --sheet Sheet1 data.xlsx
```

Returns sheet names, dimensions, cell counts, formula presence, and data ranges.

| Flag | Description |
|------|-------------|
| `--sheet <name>` | Show only the named sheet (default: all) |

### `read` — Read cell data

```bash
wb read data.xlsx
wb read --range A1:C10 data.xlsx
wb read --headers --where "Status=Failed" data.xlsx
wb read --format markdown --headers data.xlsx
wb read --all-sheets data.xlsx
wb read --limit 5 --headers data.xlsx
```

Returns stored/cached values. Use `calc` to force formula recalculation.

| Flag | Description |
|------|-------------|
| `--sheet <name>` | Read from the named sheet (default: first) |
| `--all-sheets` | Read all sheets |
| `--range <A1:D10>` | Read a specific range (default: full used range) |
| `--limit <N>` | Limit output to first N data rows |
| `--head <N>` | Alias for `--limit` |
| `--where <expr>` | Filter rows (repeatable, AND logic). Operators: `=`, `!=`, `<`, `>`, `<=`, `>=` |
| `--headers` | Treat first row as headers |
| `--include-formulas` | Include formula strings in output |
| `--include-styles` | Include style objects in output |
| `--style-summary` | Include human-readable style summary per cell |
| `--no-dates` | Disable date detection; show raw numbers |

### `edit` — Modify an existing workbook

```bash
wb edit --patch '[{"cell":"A1","value":"updated"}]' data.xlsx
echo '[{"cell":"B1","formula":"SUM(A1:A10)"}]' | wb edit data.xlsx
wb edit --dry-run --patch '[{"cell":"A1","clear":true}]' data.xlsx
```

Accepts a JSON array of patch operations via `--patch` or stdin.

| Flag | Description |
|------|-------------|
| `--patch <json>` | Patch JSON array |
| `--sheet <name>` | Default sheet for operations (default: first) |
| `--output <path>` | Save to a different file (default: overwrite input) |
| `--dry-run` | Report changes without saving |
| `--validate-only` | Validate and apply in-memory only (never saves) |
| `--atomic` | Save only if all operations succeed (default) |
| `--no-atomic` | Allow partial saves when operations fail |
| `--plan` | Include a normalized operation plan in output |

#### Patch operations

```json
[
  {"cell": "A1", "value": "hello"},
  {"cell": "B1", "value": 42},
  {"cell": "C1", "formula": "SUM(A1:B1)"},
  {"cell": "D1", "style": {"font": {"bold": true}}},
  {"cell": "A1", "clear": true},
  {"cell": "A1:C3", "clear": true},
  {"add_sheet": "NewSheet"},
  {"delete_sheet": "OldSheet"},
  {"cell": "A", "column_width": 25.0},
  {"row": 1, "row_height": 30.0}
]
```

### `create` — Create a new workbook

```bash
wb create --spec '{"sheets":["S1"],"cells":[{"cell":"A1","value":"hello"}]}' out.xlsx
echo '{"rows":[{"start":"A1","data":[["a","b"],[1,2]]}]}' | wb create out.xlsx
```

Accepts a JSON spec via `--spec` or stdin.
Unknown JSON fields are rejected.

| Flag | Description |
|------|-------------|
| `--spec <json>` | Spec JSON |

#### Spec format

```json
{
  "sheets": ["Sheet1", "Sheet2"],
  "cells": [
    {"cell": "A1", "value": "Name"},
    {"cell": "B1", "value": 42},
    {"cell": "C1", "formula": "SUM(A1:B1)"}
  ],
  "rows": [
    {"sheet": "Sheet1", "start": "A2", "data": [["Alice", 10], ["Bob", 20]]}
  ]
}
```

- `sheets` — Array of sheet names (default: `["Sheet1"]`)
- `cells` — Array of cell patch operations (same format as `edit`)
- `rows` — Row-oriented data blocks with `sheet`, `start`, and `data` fields

### `calc` — Recalculate formulas

```bash
wb calc data.xlsx
wb calc --range A1:C10 data.xlsx
wb calc --output recalculated.xlsx data.xlsx
```

Evaluates every formula and returns results (unlike `read`, which returns cached values).

| Flag | Description |
|------|-------------|
| `--sheet <name>` | Recalculate the named sheet (default: first) |
| `--range <A1:D10>` | Return results for a specific range |
| `--output <path>` | Save the recalculated workbook to a file |
| `--no-dates` | Disable date detection; show raw numbers |

### `formula list` — List supported functions

```bash
wb formula list
```

Returns all registered formula functions.

### `help` — Show help for a command

```bash
wb help read
wb help formula list
wb --mode agent help read
```

In default mode, help is human-readable text.
In `--mode agent`, help is returned as structured JSON on stdout.

### `version` — Print version

```bash
wb version
```

## Output

All commands return structured JSON by default:

```json
{
  "ok": true,
  "command": "read",
  "data": { ... },
  "meta": {
    "schema_version": "wb.v1",
    "tool_version": "dev",
    "elapsed_ms": 3
  }
}
```

Errors are written to stderr (or stdout when `--mode agent` is enabled):

```json
{
  "ok": false,
  "command": "read",
  "error": {
    "code": "FILE_NOT_FOUND",
    "message": "could not open \"missing.xlsx\": ...",
    "hint": "Check the file path. Use 'wb create' to create a new file."
  }
}
```

### Agent mode

`--mode agent` makes the CLI easier for an LLM to drive:

- Forces JSON envelopes to stdout for both success and error responses
- Returns structured help data for `wb help` and `<command> --help`
- Pairs naturally with `wb capabilities` for command discovery

Example:

```bash
wb --mode agent help read
```

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | File I/O error |
| 2 | Validation error |
| 3 | Partial failure (some operations failed) |
| 4 | Usage error |
| 99 | Internal error |
