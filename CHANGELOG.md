# Changelog

## v0.4.0

### Breaking changes

- **CLI renamed from `werkbook` to `wb`**: The CLI binary is now installed as `wb` instead of `werkbook`. Update any scripts or aliases accordingly. Install with `go install github.com/jpoz/werkbook/cmd/wb@latest`.

- **Module path changed**: The Go module path is now `github.com/jpoz/werkbook` (previously `github.com/werkbook/werkbook`). Update your `import` statements and `go.mod` accordingly.

## v0.3.0

### CLI

#### New features

- **`--limit N` / `--head N` flag for `read`**: Limit output to the first N data rows. When used with `--headers`, the header row is not counted. When combined with `--where`, the limit applies after filtering.

- **`--all-sheets` flag for `read`**: Read all sheets in a workbook in a single command. JSON output wraps sheets in a `{"sheets": [...]}` array. Markdown separates sheets with `## SheetName` headers. CSV separates sheets with `# SheetName` comment lines. Mutually exclusive with `--sheet`. All other flags (`--range`, `--limit`, `--where`, `--headers`) apply independently per sheet.

- **`--where` filter flag for `read`**: Filter rows with repeatable `--where "column<op>value"` expressions (AND semantics). Supported operators: `=`, `!=`, `<`, `>`, `<=`, `>=`. Column references use header names when `--headers` is set, or column letters (A, B, ...) otherwise. Values are compared numerically when both sides parse as numbers; otherwise compared as strings (case-insensitive for `=`/`!=`).

- **`--style-summary` flag for `read`**: Adds a human-readable `style_summary` string to each cell (e.g. `"bold, 14pt, fill:#FF0000"`). In JSON output this is a field on each cell object. In markdown/CSV output a "Style" column is appended.

#### Improvements

- **`werkbook --help` and `werkbook -h`**: Now correctly display global usage instead of erroring with "unknown command". `werkbook help` (with no subcommand) now exits with code 0 instead of 4.

- **`werkbook edit --help`**: Added a note clarifying that setting cell values does not auto-expand formula ranges.
