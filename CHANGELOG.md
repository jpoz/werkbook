# Changelog

## v0.6.1

### Improvements

- **Zero external dependencies**: Removed `spf13/cobra` and `spf13/pflag` from the module. The community library now has no external dependencies (stdlib only). The CLI commands already parsed their own flags, so cobra was a redundant routing layer replaced by a simple command dispatcher.

## v0.6.0

### New formula functions (55 functions)

#### Complex number functions (22)
- `COMPLEX`, `IMAGINARY`, `IMREAL`, `IMABS`
- `IMSUM`, `IMSUB`, `IMPRODUCT`, `IMDIV`
- `IMARGUMENT`, `IMCONJUGATE`, `IMSQRT`, `IMPOWER`
- `IMEXP`, `IMLN`, `IMLOG2`, `IMLOG10`
- `IMSIN`, `IMCOS`, `IMTAN`
- `IMSINH`, `IMCOSH`, `IMSECH`, `IMCSC`, `IMCOT`, `IMCSCH`

#### Statistical distribution functions (25)
- `BINOM.DIST`, `BINOM.INV`, `NEGBINOM.DIST`
- `POISSON.DIST`, `EXPON.DIST`, `WEIBULL.DIST`
- `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`
- `LOGNORM.DIST`, `LOGNORM.INV`
- `T.DIST`, `T.INV`, `T.DIST.RT`, `T.DIST.2T`, `T.INV.2T`
- `CHISQ.DIST`, `CHISQ.INV`, `CHISQ.DIST.RT`, `CHISQ.INV.RT`
- `F.DIST`, `F.INV`, `F.DIST.RT`, `F.INV.RT`
- `HYPGEOM.DIST`

#### Other new functions (8)
- `BETA.DIST`, `BETA.INV`
- `CONFIDENCE.NORM`, `CONFIDENCE.T`
- `PHI`, `GAUSS`, `SKEW.P`
- `ISPMT`
- `UNICHAR`, `UNICODE`
- `XMATCH`, `SORTBY`
- `CHOOSECOLS`, `CHOOSEROWS`
- `PERCENTILE.INC`, `QUARTILE.INC` (aliases)

### Features

- **Lambda support**: Formula engine now supports `LAMBDA` functions and 1904 date system
- **Dynamic array formulas**: Added support for dynamic array formula spill ranges
- **Sheet copying and cloning**: New `Sheet.Copy()` and `File.CloneSheet()` methods
- **Merge cell support**: Added sheet operations for merged cells
- **Defined names**: Added defined name resolution API and operations
- **Table support**: Added table and calc properties support
- **`_xlfn` prefix handling**: Extended function prefix mapping for better spreadsheet compatibility

### CLI improvements

- **`--show-formulas` flag**: Display raw formulas instead of computed values
- **Human-readable text output**: New default text output format with markdown table formatting

### Bug fixes

- Improved `NORM.INV` precision with Newton-Raphson refinement
- Fixed `GAMMA.DIST`, `LOGNORM.DIST`/`LOGNORM.INV`, and `WEIBULL.DIST` to match spreadsheet behavior
- Fixed `CODE` function to return underscore (95) for unmappable characters
- Fixed tilde escape handling in `COUNTIF`/`SUMIF` wildcard criteria
- Fixed `INDEX` to return `#VALUE!` for `row_num=0` in non-array context
- Fixed `_xludf.` prefixed functions to return `#NAME?` matching spreadsheet behavior
- Replaced `fmt.Sprintf` with `strconv` for number formatting (performance)

### Test coverage

- Added comprehensive tests for 70+ formula functions

## v0.4.1

### Bug fixes

- **`wb dep`: skip empty cells in range expansion**: When a formula depends on a range (e.g. `=SUM(A1:A100)`), the dep command no longer lists cells that have no value and no formula, reducing noise in the output.

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
