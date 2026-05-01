# Changelog

## v0.11.0

### CLI

- **`wb serve` command**: New NDJSON-over-stdin/stdout interface for stateful,
  long-running workbook sessions. Avoids the per-command cost of opening and
  saving a file, letting external agents perform multiple operations on a
  single in-memory workbook.
- **`--dry-run` / `--validate-only` flags on `create`**: Build workbooks in
  memory without disk I/O so callers can verify complex specifications
  independently of filesystem side effects.
- **Embedded input JSON schemas in `capabilities`**: Capabilities output now
  includes formal JSON schemas for every structured input, giving agents a
  machine-readable contract to validate payloads against.
- **Formula discovery metadata**: Formula listings expose argument counts and
  types so callers can introspect what each function accepts.
- **Typed date/time cell support**: Cell operations accept an optional `type`
  field for `date`, `datetime`, and `time` values. Strings are parsed and
  converted to Excel serials with appropriate default number formats applied,
  preventing Excel from displaying raw serial numbers.
- **Inline typed row values**: Row data elements accept `{type, value, formula,
  style}` objects alongside scalar entries.
- **Duplicate-write detection**: `wb` warns when multiple operations target the
  same cell (e.g. a later row write clobbering an earlier formula).

### Compatibility

- **Cached spill values for dynamic array followers**: Saved files now persist
  cached values for spill follower cells and emit `<f ca="1"/>` tags. Excel
  re-evaluates on load and doesn't need this, but Apple Numbers and other
  consumers that cannot evaluate `_xlfn._xlws.FILTER` and friends now display
  spilled results correctly.

### Formula engine

- **`TRANSPOSE` registered as a dynamic array function**: The engine now
  correctly identifies and spills `TRANSPOSE` results.
- **Date-text operands in criteria**: `COUNTIF` / `SUMIF` / `AVERAGEIF`
  families coerce date-formatted text operands like `">=2026-01-01"` to
  numeric serials before comparison, matching Excel. ISO dash, US slash, and
  ISO slash formats are recognized.
- **`INDEX(range, 0)` implicit intersection**: In non-array context,
  `INDEX(range, 0)` now performs implicit intersection at the formula cell's
  row instead of returning `#VALUE!`.
- **`REGEXEXTRACT` mode 1 layout**: Returns a horizontal `1×n` array instead
  of a vertical `n×1` array, matching Excel.
- **Implicit intersection for range-origin arrays**: Array results that carry
  a range origin now collapse to a single cell correctly when used in
  non-array context (e.g. `INDEX` returning a vector).
- **Argument validation alignment**: `XLOOKUP`, `MATCH`, `TAKE`, and `DROP`
  argument validation matches Excel for omitted parameters and dimension
  constraints.
- **`COUNTA` / `SORT`**: Empty values and errors are processed correctly
  across argument types.

## v0.10.0

### Formula Engine v2

A large internal restructuring of the formula evaluator. Public APIs remain
compatible; the migration separates scalars, computed arrays, and worksheet
references into distinct runtime types so dynamic arrays, spill behavior, and
range semantics can be reasoned about explicitly.

- **`EvalValue` runtime model**: New internal envelope with `EvalScalar`,
  `EvalArray`, `EvalRef`, and `EvalKindError` kinds. Worksheet ranges and
  computed arrays no longer share a single `ValueArray` representation.
- **`RefValue` with bounds and 3D support**: Reference identity is preserved
  through `OFFSET`, `INDIRECT`, and selector functions instead of being
  flattened into anonymous arrays.
- **`ArrayValue` with explicit shape and `SpillClass`**: Replaces broad
  reliance on the legacy `NoSpill` flag with `SpillNone` / `SpillBounded` /
  `SpillUnbounded` / `SpillScalarOnly`.
- **Lazy `Grid` iteration**: Reducers and scans iterate references and arrays
  without eagerly materializing full-column / full-row inputs.
- **Unified `FuncSpec` contracts**: Argument loading, adaptation, and array
  lifting policy live on each function's spec instead of in scattered global
  maps. The old `funcMetaByName` semantic map has been removed; `RegisterWithMeta`
  is now a compatibility shim that produces a `FuncSpec`.
- **`CellEvalOutcome` (raw / display / spill split)**: The workbook layer now
  tracks raw evaluation result, displayed anchor value, and spill plan
  separately, so blocked spills cleanly expose `#SPILL!` while preserving the
  underlying array.
- **Eval-aware criteria & reducer paths**: `COUNTIF` / `SUMIF` / `AVERAGEIF`
  families and `SUM` / `COUNT` / `AVERAGE` / `MIN` / `MAX` consume `EvalValue`
  and `Grid` directly with far less adapter churn.

### New formula functions

- **Regex (3)**: `REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE` (Excel 2024).
  Cached pattern compiler, Excel-style `$N` replacement expansion, and
  `instance_num` support. Patterns requiring backreferences/lookarounds (not
  supported by Go's RE2) return `#VALUE!`.

### Features

- **Intersection & union operators**: Space-separated range intersection
  (`A1:C3 B2:D4`) returns the rectangular overlap or `#NULL!`; parenthesized
  union references (`(A1:A2,C1:C2)`) flatten constituent areas.
- **Dynamic range references**: `A1:INDEX(A:A,n)` and similar reference-
  producing function calls are accepted on either side of `:` and built at
  runtime via `OpBuildRange`.
- **LAMBDA improvements**: `CellRef` accepts bare identifiers as parameter
  names (e.g. `running_sum`); `_xlpm.` prefixes are stripped on serialization;
  workbook-level named LAMBDAs can be invoked like functions; `ISOMITTED` is
  supported.
- **Sheet-qualified full-row ranges**: `Sheet1!2:3` and `'venture-dist'!2:3`
  now parse, mirroring the long-supported column-only form.

### Bug fixes

- **Error classification**: `evaluateFormula` no longer collapses every
  failure to `#NAME?`. Parse/compile failures produce `#NAME?`; expansion
  overflow and runtime engine errors produce `#VALUE!`.
- **Quoted sheet names in rewriters**: `ExpandTableRefs`, defined name
  expansion, `r1c1ToA1`, and `INDIRECT`'s sheet-prefix extraction are now
  quote-aware and correctly handle escaped `''` within sheet names.
- **Implicit intersection for mixed operands**: When one operand is a
  range-derived scalar and the other an anonymous array, the anonymous array
  is collapsed to its top-left element, matching Excel.
- **`IFERROR`/`IFNA` under legacy array context**: Nested under
  `SUMPRODUCT`-style array forcers, these apply implicit intersection;
  dynamic-array natives (`FILTER`, `SORT`, `UNIQUE`, …) still lift arrays.
- **`SUMIFS` error propagation**: Errors from referenced cells now propagate
  instead of being silently ignored.
- **`BYCOL`/`BYROW`**: Return `#VALUE!` (not `#CALC!`) when the lambda yields
  an array.
- **Element-wise lifting**: Financial scalar functions `DB`, `DDB`, `SLN`,
  `SYD` are registered as element-wise.

## v0.9.3

### Bug fixes

- **VLOOKUP/HLOOKUP approximate match**: Switched approximate match (`range_lookup=TRUE`) from linear scan to binary search, matching Excel's behavior on sorted data
- **XLOOKUP multi-column return arrays**: Return-array orientation logic now distinguishes row- vs column-oriented lookups, so multi-column returns select the matching column instead of flattening all return values

## v0.9.2

### Bug fixes

- **Tolerate leading `=` in `SetFormula`**: Formulas passed with a leading `=` (as typed in Excel) are now stripped at the public-API boundary. Previously the `=` nested inside the OOXML `<f>` element and saved as a `#NAME?` error cell.

## v0.9.1

### Features

- **Full column/row ranges in defined names**: `ResolveDefinedName` now supports full column (`A:H`) and full row (`2:2`) references, routing through the internal range resolver so dynamic spills and sparse ranges are preserved
- **`WithoutFormulas` open option**: New option on `Open` functions that skips dependency graph construction and formula evaluation for metadata-only reads, speeding up `wb info` and similar workflows
- **Spill array performance**: Precomputed bounds cache replaces O(rows) scans in hot evaluation loops, plus a new spill overlay tracking system that manages array formula spill states without full sheet iteration

### Bug fixes

- **DATEVALUE in FILTER**: DATEVALUE now participates in inherited array context, fixing FILTER operations with full-column DATEVALUE expressions
- **Uncached dynamic array spills**: Dynamic array formulas without cached spill ranges now evaluate on file load so dependent formulas see the spilled results
- **SUMPRODUCT / IF array context**: IF no longer inherits array context from SUMPRODUCT, matching Excel behavior
- **Oversized defined name ranges**: Bounded defined name ranges now reject materialization beyond the range cell count limit instead of risking OOM

## v0.9.0

### Features

- **Shared formula support**: Added shared formula parsing and external reference filtering in OOXML import
- **Dynamic array spill conflict detection**: Detect and handle spill conflicts for dynamic array formulas, matching Excel behavior
- **Spill control for imported formulas**: Dynamic array spill behavior is now preserved through import
- **Check command**: New `wb check` command for validating engine parity against Excel
- **Check config**: Configurable check runs with formula alignment improvements
- **Exported column conversion functions**: `ColumnToIndex` and `IndexToColumn` are now public API

### Bug fixes

- **Formula formatting and math alignment**: Aligned number formatting and math operations with Excel behavior
- **Inherited array context in error wrappers**: Fixed array context propagation through error-handling formula wrappers
- **IF in inherited array context**: IF function now correctly participates in inherited array context

## v0.8.1

- Added MIT license

## v0.8.0

### New formula functions (9 functions)

#### Higher-order array functions (6)
- `MAP`, `REDUCE`, `SCAN`
- `BYROW`, `BYCOL`, `MAKEARRAY`

#### Information functions (3)
- `SHEET`, `SHEETS`, `AREAS`

### Features

- **Dynamic array roundtrip**: Preserve dynamic array metadata (spill ranges, formulas) through read/write cycles
- **Spill array aggregation**: Range-based functions (SUM, AVERAGE, etc.) now correctly include spill array values

### Bug fixes

- **Dynamic array serialization**: Simplified and corrected dynamic array formula serialization in OOXML output
- **RANDARRAY spill handling**: Fixed spill range evaluation for RANDARRAY

### Test coverage

- Added comprehensive tests for 50+ formula functions covering financial, statistical, lookup, text, date, math, and database categories

## v0.7.1

### Security

- **Expansion bomb prevention**: Added size limits to formula range expansion to prevent denial-of-service via crafted spreadsheets

### Bug fixes

- **Array/scalar lifting**: Scalar functions now correctly lift over array arguments, matching Excel behavior
- **Full-column ranges**: Fixed AGGREGATE and other functions handling of full-column ranges (e.g. `A:A`)
- **Range trimming**: Preserved logical bounds when trimming ranges, fixing edge cases where trimmed ranges lost their original dimensions
- **Argument evaluation modes**: Refined how formula function arguments are evaluated, improving correctness for functions that mix scalar and range parameters

### Other

- Added regression tests for whole-column recalculation
- Updated README description

## v0.7.0

### New formula functions (64 functions)

#### Financial functions (25)
- `ACCRINT`, `ACCRINTM`
- `AMORDEGRC`, `AMORLINC`
- `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`
- `DISC`, `INTRATE`, `RECEIVED`
- `DURATION`, `MDURATION`
- `FVSCHEDULE`
- `PRICE`, `YIELD`, `PRICEDISC`, `YIELDDISC`, `PRICEMAT`, `YIELDMAT`
- `TBILLPRICE`, `TBILLYIELD`, `TBILLEQ`

#### Statistical functions (10)
- `AGGREGATE`
- `BINOM.DIST.RANGE`, `PROB`
- `CHISQ.TEST`, `F.TEST`, `T.TEST`, `Z.TEST`
- `MODE.MULT`, `STDEVA`, `STDEVPA`

#### Database functions (12)
- `DSUM`, `DAVERAGE`, `DCOUNT`, `DCOUNTA`
- `DGET`, `DMAX`, `DMIN`, `DPRODUCT`
- `DSTDEV`, `DSTDEVP`, `DVAR`, `DVARP`

#### Regression and array functions (6)
- `LINEST`, `LOGEST`, `TREND`, `GROWTH`
- `EXPAND`, `RANDARRAY`

#### Engineering functions (5)
- `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`
- `IMSEC`

#### Other new functions (6)
- `ENCODEURL`, `HYPERLINK`
- `ISO.CEILING`
- `ISREF`
- `LET`, `OFFSET`

### Features

- **Core properties support**: Read and write workbook core properties (title, author, description, etc.)
- **Range guards**: Bounds checking for range operations
- **Qualified local names**: Formula engine supports qualified local defined name references

### Bug fixes

- **VLOOKUP**: Exact match now supports wildcards and skips empty cells
- **HLOOKUP**: Added wildcard matching; returns `#VALUE!` for `row_index < 1`
- **XLOOKUP**: Fixed `search_mode`, binary search support, and omitted `if_not_found` handling
- **SUMPRODUCT**: Treats boolean and text cell values as 0 to match Excel
- **SWITCH**, **LEN**, **UNICHAR**: Aligned behavior with Excel
- **OFFSET**: Match Excel behavior for negative height/width and nesting
- **T()**: Propagates error values instead of converting to string
- **DAY**, **MONTH**, **YEAR**: Accept date strings like Excel
- **WEEKDAY**: Match Excel behavior for serial numbers <= 60
- **PV**: Returns `#NUM!` instead of `Inf` for degenerate inputs
- **FVSCHEDULE**: Rejects booleans in schedule to match Excel
- **SYD**: Fixed argument validation and evaluation
- **AMORDEGRC**: Fixed rounding to match Excel behavior
- **AMORLINC**: Match Excel behavior
- **LINEST/LOGEST**: Fixed F-statistic calculation for near-perfect and perfect fit data
- **NORM.S.INV**: Improved precision with Newton-Raphson refinement
- **NORM.S.DIST**: Fixed spec/xlsx desync for 1-arg error case
- **PROB**: Match Excel behavior for zero and negative probabilities
- **T.TEST**: Paired mode skips non-numeric pairs to match Excel
- **BESSELI/BESSELJ**: Improved precision to match Excel
- **TBILLEQ**: Uses semi-annual compounding for DSM > 182 days
- **DISC/INTRATE/RECEIVED**: Fixed basis=1 day count for cross-year periods
- **COUP/DURATION** functions: Match Excel behavior
- **ISREF**: Returns `TRUE` for `INDIRECT`/`OFFSET` ref-returning functions
- **collectNumeric**: Ignores text/bool/empty cell refs (affects STDEV.S and other statistical functions)

### Test coverage

- Added comprehensive tests for 80+ formula functions

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
