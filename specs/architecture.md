# Werkbook Architecture Specification

## Overview

Werkbook is a Go library for reading, writing, and recalculating Excel XLSX files. It prioritizes fast recalculation through a compiled formula engine with a proper dependency graph, while maintaining full compatibility with the Office Open XML (OOXML) standard.

### Design Principles

1. **Fast recalculation** - Formulas are parsed once, compiled to bytecode, and evaluated via a stack-based VM. A dependency graph enables incremental recalculation.
2. **Correct by default** - Match Excel behavior for formula evaluation, type coercion, and edge cases.
3. **Memory efficient** - Sparse cell storage; streaming and lazy-loading planned for future phases.
4. **Simple API** - Provide a clean, idiomatic Go API that covers the common 90% of use cases without exposing OOXML internals.

---

## Package Structure

```
werkbook/
  werkbook.go          # File struct, Open, New, SaveAs, Recalculate
  sheet.go             # Sheet type and operations
  cell.go              # Cell type, getting/setting values
  row.go               # Row iteration and access
  value.go             # Value tagged union (TypeEmpty, TypeNumber, TypeString, TypeBool, TypeError)
  coords.go            # A1-style coordinate system (CellNameToCoordinates, etc.)
  date.go              # Date handling, Excel serial number conversion
  errors.go            # Error type definitions
  formula/
    lexer.go           # Formula tokenizer
    token.go           # Token types
    parser.go          # Token stream -> AST (Pratt parser)
    ast.go             # AST node types
    cellref.go         # Cell/range reference parsing
    compiler.go        # AST -> bytecode (99 registered functions)
    opcodes.go         # Bytecode instruction set
    eval.go            # Stack-based VM evaluator
    types.go           # Value types for formula engine (Number, String, Bool, Error, Array)
    depgraph.go        # Cell dependency graph for incremental recalculation
    functions_math.go  # Math function implementations (27 functions)
    functions_text.go  # Text function implementations (22 functions)
    functions_stat.go  # Statistical function implementations (16 functions)
    functions_date.go  # Date/time function implementations (11 functions)
    functions_lookup.go # Lookup function implementations (7 functions)
    functions_logic.go # Logical function implementations (7 functions)
    functions_info.go  # Information function implementations (11 functions)
  ooxml/
    reader.go          # ZIP/XML reading
    writer.go          # ZIP/XML writing
    workbook.go        # Intermediate data structures (WorkbookData, SheetData, etc.)
    worksheet.go       # Worksheet XML handling
    styles.go          # Style definitions (minimal default only)
    sharedstrings.go   # Shared string table
    relationships.go   # Package relationships
    contenttypes.go    # Content type mappings
  _examples/
    create/main.go     # Example: create workbook with formulas
    calculate/main.go  # Example: formula engine demo (amortization, grades, etc.)
```

**Not yet implemented** (planned for future phases):
- `style.go` — Style definitions and application (user-facing API)
- `formula/functions_fin.go` — Financial function implementations
- `formula/functions_eng.go` — Engineering function implementations
- `ooxml/calcchain.go` — Calculation chain XML
- `stream/` — Streaming reader/writer for large files

---

## Test Coverage Summary

| Package | Coverage | Test Files | Notes |
|---------|----------|------------|-------|
| `werkbook` (root) | 77.2% | 12 | Round-trip, iteration, formulas, caching, dep graph, multi-sheet, values |
| `werkbook/formula` | 82.9% | 13 | Lexer, parser, compiler, eval, dep graph, all 7 function categories |
| `werkbook/ooxml` | 0.0% | 0 | Exercised indirectly via integration tests |
| **Overall** | ~**77%** | **25** | No benchmarks yet |

---

## Core Types

### File

The root object representing an open workbook.

```go
type File struct {
    sheets     []*Sheet
    sheetNames []string
    calcGen    uint64              // incremented on any cell mutation; starts at 1
    evaluating map[cellKey]bool    // tracks cells being evaluated (circular ref detection)
    deps       *formula.DepGraph   // cell dependency graph for incremental recalculation
}

// cellKey identifies a cell across the entire workbook for circular ref detection.
type cellKey struct {
    sheet string
    col   int
    row   int
}
```

`New()` accepts functional options (currently only `FirstSheet(name)` to set the initial sheet name). `Open()` reads an XLSX, parses all sheets eagerly, compiles all formulas, and registers them in the dependency graph.

### Sheet

```go
type Sheet struct {
    file *File
    name string
    rows map[int]*Row   // sparse row storage, keyed by 1-based row number
}
```

### Row and Cell

```go
type Row struct {
    sheet  *Sheet
    num    int             // 1-based row number
    cells  map[int]*Cell   // sparse cell storage, keyed by 1-based column
}

type Cell struct {
    col       int                        // 1-based column number
    value     Value                      // current value (literal or last computed)
    formula   string                     // formula text (e.g., "SUM(A1:A10)"), "" if none
    compiled  *formula.CompiledFormula   // cached compiled bytecode
    cachedGen uint64                     // file.calcGen when value was last computed from formula
    dirty     bool                       // flagged by dependency graph on upstream mutation
}
```

Formula cells store both the formula text and its compiled bytecode. The `cachedGen` field enables lazy evaluation: when `GetValue` is called on a formula cell whose `cachedGen < file.calcGen` or whose `dirty` flag is set, the formula is re-evaluated via the VM.

### Value (Public API — `value.go`)

A tagged union for cell values in the public API. Error cells store the error string (e.g., `"#DIV/0!"`) in the `String` field with `Type: TypeError`.

```go
type Value struct {
    Type   ValueType
    Number float64
    String string
    Bool   bool
}

type ValueType int

const (
    TypeEmpty ValueType = iota
    TypeNumber
    TypeString
    TypeBool
    TypeError
)
```

### Value (Formula Engine — `formula/types.go`)

The formula engine uses a separate `Value` type with a numeric error code and array support for range results.

```go
type Value struct {
    Type  ValueType
    Num   float64
    Str   string
    Bool  bool
    Err   ErrorValue
    Array [][]Value  // used by ValueArray for range results
}

type ValueType byte

const (
    ValueEmpty  ValueType = iota
    ValueNumber
    ValueString
    ValueBool
    ValueError
    ValueArray  // for range results in VM
)

type ErrorValue byte

const (
    ErrValDIV0        ErrorValue = iota // #DIV/0!
    ErrValNA                            // #N/A
    ErrValNAME                          // #NAME?
    ErrValNULL                          // #NULL!
    ErrValNUM                           // #NUM!
    ErrValREF                           // #REF!
    ErrValVALUE                         // #VALUE!
    ErrValSPILL                         // #SPILL!
    ErrValCALC                          // #CALC!
    ErrValGETTINGDATA                   // #GETTING_DATA
)
```

Conversion between the two `Value` types happens at the boundary in `sheet.go` (`formulaValueToValue` / `valueToFormulaValue`).

---

## Formula Engine (the key differentiator)

The formula engine is a key differentiator. Werkbook compiles formulas once to bytecode and evaluates them in a tight VM loop, rather than re-tokenizing on every evaluation or dispatching functions via reflection.

### Pipeline

```
Formula String
    |
    v
[Lexer] -----> Token Stream
    |
    v
[Parser] ----> AST (expression tree)
    |
    v
[Compiler] --> Bytecode ([]Instruction)
    |
    v
[VM] ---------> Value (result)
```

### Phase 1: Lexer (`formula/lexer.go`)

Tokenizes Excel formula strings into a flat token stream. Handles:

- Cell references: `A1`, `$A$1`, `Sheet1!A1`, `'Sheet Name'!A1`
- Ranges: `A1:B5`, `A:A`, `1:1`
- Numbers: `123`, `1.5`, `1.5E10`
- Strings: `"hello"`
- Booleans: `TRUE`, `FALSE`
- Errors: `#N/A`, `#DIV/0!`
- Operators: `+`, `-`, `*`, `/`, `^`, `&`, `=`, `<>`, `<`, `>`, `<=`, `>=`, `%`
- Functions: `SUM(`, `IF(`
- Parentheses, commas, semicolons
- Array literals: `{1,2;3,4}`

```go
type TokenType byte

const (
    TokNumber TokenType = iota
    TokString
    TokBool
    TokError
    TokCellRef
    TokRange
    TokFunc
    TokOp
    TokLParen
    TokRParen
    TokComma
    TokArrayOpen   // {
    TokArrayClose  // }
    TokSemicolon   // row separator in array literals
    TokPercent
    TokEOF
)

type Token struct {
    Type  TokenType
    Value string // raw text
    Pos   int    // position in source for error reporting
}
```

### Phase 2: Parser (`formula/parser.go`)

Pratt parser (top-down operator precedence) that produces an AST. Pratt parsing handles prefix/infix/postfix operators cleanly and is easier to extend than shunting-yard.

```go
type Node interface{ node() }

type NumberLit struct { Value float64 }
type StringLit struct { Value string }
type BoolLit   struct { Value bool }
type ErrorLit  struct { Code  ErrorCode }

type CellRef struct {
    Sheet    string // "" for same-sheet
    Col      int
    Row      int
    AbsCol   bool
    AbsRow   bool
}

type RangeRef struct {
    From CellRef
    To   CellRef
}

type UnaryExpr struct {
    Op      string // "-", "+"
    Operand Node
}

type BinaryExpr struct {
    Op    string // "+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="
    Left  Node
    Right Node
}

type PostfixExpr struct {
    Op      string // "%"
    Operand Node
}

type FuncCall struct {
    Name string
    Args []Node
}

type ArrayLit struct {
    Rows [][]Node
}
```

### Phase 3: Compiler (`formula/compiler.go`)

Walks the AST and emits bytecode instructions. The instruction set is intentionally small.

```go
type OpCode byte

const (
    OpPushNum    OpCode = iota // push float64 constant
    OpPushStr                  // push string constant
    OpPushBool                 // push bool constant
    OpPushError                // push error constant
    OpPushEmpty                // push empty value

    OpLoadCell                 // push value of cell ref (operand: sheet, col, row)
    OpLoadRange                // push array from range (operand: sheet, col1, row1, col2, row2)

    OpAdd
    OpSub
    OpMul
    OpDiv
    OpPow
    OpNeg                      // unary negate
    OpPercent                  // divide by 100
    OpConcat                   // string concatenation (&)

    OpEq
    OpNe
    OpLt
    OpLe
    OpGt
    OpGe

    OpCall                     // call function (operand: func ID, arg count)

    OpMakeArray                // construct array from stack values
)

type Instruction struct {
    Op      OpCode
    Operand uint32 // meaning depends on opcode
}

type CompiledFormula struct {
    Source  string          // original formula text (for serialization)
    Code    []Instruction   // bytecode
    Consts  []Value         // constant pool (numbers, strings)
    Refs    []CellAddr      // cell references used (for dependency tracking)
    Ranges  []RangeAddr     // range references used (for dependency tracking)
}
```

The `Refs` and `Ranges` fields are populated during compilation and are used to build the dependency graph without re-parsing.

### Phase 4: VM (`formula/eval.go`)

A stack-based evaluator that executes compiled formulas. Cell and range lookups are abstracted via the `CellResolver` interface, keeping the VM decoupled from `Sheet`/`File`.

```go
// CellResolver abstracts cell/range lookups so the VM has no dependency on Sheet.
type CellResolver interface {
    GetCellValue(addr CellAddr) Value
    GetRangeValues(addr RangeAddr) [][]Value
}

// EvalContext provides context about the current evaluation environment.
type EvalContext struct {
    CurrentCol   int
    CurrentRow   int
    CurrentSheet string
}

// Eval executes a compiled formula and returns the result.
func Eval(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext) (Value, error) {
    stack := make([]Value, 0, 16)
    for _, inst := range cf.Code {
        switch inst.Op {
        case OpPushNum:
            push(cf.Consts[inst.Operand])
        case OpLoadCell:
            push(resolver.GetCellValue(cf.Refs[inst.Operand]))
        case OpLoadRange:
            rows := resolver.GetRangeValues(cf.Ranges[inst.Operand])
            push(Value{Type: ValueArray, Array: rows})
        case OpAdd:
            // ... type coercion + arithmetic
        case OpCall:
            funcID := inst.Operand >> 8
            argc := inst.Operand & 0xFF
            // dispatch via callFunction() with function pointer lookup
        // ... other opcodes
        }
    }
    return stack[0], nil
}
```

Key performance properties:
- **No reflection** - function dispatch via `callFunction()` with a switch over function IDs
- **No re-parsing** - bytecode is compiled once per formula and cached in `Cell.compiled`
- **Cache-friendly** - instructions are a compact byte stream; the operand stack is a contiguous slice

The `fileResolver` type in `sheet.go` implements `CellResolver`, providing cross-sheet cell lookups with lazy formula evaluation (triggering `resolveCell` on referenced formula cells).

### Function Registry (`formula/compiler.go` + `formula/eval.go`)

Functions are identified at compile time by index into a sorted `knownFunctions` array (99 entries). At eval time, `callFunction()` dispatches via a switch on function ID. No reflection.

```go
// Compile-time: sorted array of known function names
var knownFunctions = [...]string{"ABS", "ACOS", "AND", ..., "YEAR"} // 99 entries
var funcNameToID map[string]int // populated at init

// Eval-time: dispatch by function ID
func callFunction(funcID int, args []Value, ctx *EvalContext) Value {
    switch funcID {
    case funcNameToID["SUM"]:
        return fnSUM(args)
    case funcNameToID["IF"]:
        return fnIF(args)
    // ...
    }
}
```

---

## Dependency Graph (`formula/depgraph.go`) ✅

Werkbook tracks which cells depend on which other cells and only recalculates what is necessary, rather than clearing all caches on any mutation.

### Data Structure

```go
// QualifiedCell is a fully-qualified cell address (sheet name is never empty).
type QualifiedCell struct {
    Sheet string
    Col   int // 1-based
    Row   int // 1-based
}

type DepGraph struct {
    // Forward edges: formula cell → set of cells it reads
    dependsOn  map[QualifiedCell]map[QualifiedCell]bool
    // Reverse edges: data cell → set of formula cells that read it
    dependents map[QualifiedCell]map[QualifiedCell]bool
    // Range subscriptions for containment checks
    rangeSubs  []rangeSub
}

type rangeSub struct {
    formulaCell QualifiedCell
    rng         RangeAddr // always fully qualified
}
```

### Operations

**Register a formula:**
When `SetFormula` is called, the formula is compiled, and `Refs`/`Ranges` from the `CompiledFormula` are used to register edges. Unqualified refs (same-sheet) are resolved to the owning sheet name.

```go
func (g *DepGraph) Register(formulaCell QualifiedCell, owningSheet string, refs []CellAddr, ranges []RangeAddr)
```

**Unregister a formula:**
Removes all forward/reverse edges and range subscriptions for a cell. Called automatically before re-registering and when deleting sheets.

```go
func (g *DepGraph) Unregister(formulaCell QualifiedCell)
```

**Query dependents:**
`DirectDependents` returns immediate dependents (point refs + range containment). `Dependents` performs a BFS to return all transitive dependents.

```go
func (g *DepGraph) DirectDependents(cell QualifiedCell) []QualifiedCell
func (g *DepGraph) Dependents(changed QualifiedCell) []QualifiedCell
```

### Recalculation Flow

```
1. User calls sheet.SetValue("A1", 42)
2.   -> cell.value = 42
3.   -> file.calcGen++ (global generation counter)
4.   -> file.invalidateDependents("Sheet1", col, row)
5.       -> deps.Dependents(A1) returns [B1, C1, D5] (BFS transitive closure)
6.       -> for each dependent: cell.dirty = true
7. User calls sheet.GetValue("D5") or file.Recalculate()
8.   -> if cell.dirty || cell.cachedGen < file.calcGen:
9.        vm.Eval(cell.compiled, ...) -> new value
10.       cell.value = new value
11.       cell.cachedGen = file.calcGen
12.       cell.dirty = false
```

This means:
- Changing a cell value does NOT trigger recalculation (just marks dirty cells via BFS).
- Reading a formula cell's value triggers lazy evaluation if dirty or stale.
- `Recalculate()` eagerly evaluates all dirty cells.
- Only the affected subgraph is recalculated, not the entire workbook.

### Circular Reference Handling

Circular references are detected at evaluation time via the `evaluating` map on `File`. When a cell attempts to evaluate itself recursively, the cycle is detected and an error value is returned. Iterative convergence (MaxIterations/IterationTolerance) is not yet implemented.

### Not Yet Implemented

- **Topological sort for recalc order** — `Recalculate()` currently iterates all cells rather than using Kahn's algorithm on the dirty set.
- **Iterative circular ref resolution** — Cycles are detected but not iteratively converged.
- **Thread safety** — No mutex; the graph is not safe for concurrent access.

---

## Reading XLSX Files (`ooxml/reader.go`)

### Process

1. Open ZIP archive via `zip.NewReader`
2. Parse `_rels/.rels` for package relationships
3. Parse `xl/workbook.xml` -> sheet list and ordering
4. Resolve sheet file paths via `xl/_rels/workbook.xml.rels`
5. Parse `xl/sharedStrings.xml` -> string table (supports rich text via `<r>` elements)
6. For each sheet, parse `xl/worksheets/sheet{N}.xml` -> rows and cells
7. Resolve SST indices (cells with `Type: "s"`) to actual strings

All sheets are parsed eagerly during `Open()`. After `ReadWorkbook` returns a `WorkbookData`, `fileFromData()` converts it to the public `File`/`Sheet`/`Cell` structures, compiles all formulas, and registers them in the dependency graph.

### Not Yet Implemented

- Lazy sheet loading (parse on first access)
- Streaming reader for large files
- Styles parsing (only default stylesheet written)
- Calculation chain (`xl/calcChain.xml`) parsing

---

## Writing XLSX Files (`ooxml/writer.go`)

### Process

1. `File.buildWorkbookData()` converts `Sheet` -> `SheetData` (sorting rows/cells, converting values to OOXML format)
2. `WriteWorkbook()` creates a ZIP writer and serializes:
   - `[Content_Types].xml`
   - `_rels/.rels` (package relationships)
   - `xl/workbook.xml` (sheet list)
   - `xl/_rels/workbook.xml.rels` (sheet relationships)
   - `xl/styles.xml` (hardcoded default)
   - `xl/worksheets/sheet{N}.xml` for each sheet
   - `xl/sharedStrings.xml` (SST built during write, deduplicating strings; cells mutated to SST indices)
3. Close ZIP

### Not Yet Implemented

- Streaming writer for large files
- Calculation chain serialization (`xl/calcChain.xml`)
- Rich style serialization (beyond default stylesheet)

---

## API Surface

### Opening and Creating

```go
// Open an existing file
f, err := werkbook.Open("report.xlsx")

// Create a new file (default first sheet: "Sheet1")
f := werkbook.New()
f := werkbook.New(werkbook.FirstSheet("Data"))

// Save
err := f.SaveAs("output.xlsx")
```

### Sheet Operations

```go
sheet := f.Sheet("Sheet1")        // returns nil if not found
sheet, err := f.NewSheet("Data")  // error if name already exists
err := f.DeleteSheet("Sheet2")    // error if only sheet
names := f.SheetNames()
```

### Cell Access

```go
// Get/set values
val, err := sheet.GetValue("A1")        // returns Value (lazy-evaluates formulas)
err := sheet.SetValue("A1", 42)         // sets number
err := sheet.SetValue("A1", "hello")    // sets string
err := sheet.SetValue("A1", true)       // sets bool
err := sheet.SetValue("A1", time.Now()) // sets date (stored as Excel serial number)

// Formulas (without leading '=')
err := sheet.SetFormula("B1", "SUM(A1:A100)")
formula, err := sheet.GetFormula("B1")  // returns "SUM(A1:A100)"

// GetValue on a formula cell triggers lazy evaluation if dirty/stale
result, err := sheet.GetValue("B1")     // returns computed Value

// Bulk read (Go 1.23 iterators)
for row := range sheet.Rows() {
    for _, cell := range row.Cells() {
        cell.Col()     // 1-based column
        cell.Value()   // current Value
        cell.Formula() // formula text or ""
    }
}

// Sheet metadata
sheet.MaxRow()  // highest row with data
sheet.MaxCol()  // highest column with data

// Debug output
sheet.PrintTo(os.Stdout) // tabular text dump
```

### Recalculation

```go
// Recalculate all dirty/stale formula cells eagerly
f.Recalculate()
```

### Styles (not yet implemented)

Only a hardcoded default stylesheet is written (Calibri 11pt). A user-facing styles API is planned for a future phase.

---

## Performance Design

| Operation | Naive Approach | Werkbook |
|-----------|----------------|----------|
| Parse formula | Every evaluation (re-tokenize) | Once (compile to bytecode, cache in `Cell.compiled`) |
| Function dispatch | Reflection-based | Switch on function ID (no reflection) |
| Dependency tracking | None (clear all caches on any mutation) | DAG with BFS incremental invalidation |
| Recalc scope | All formula cells | Only dirty cells + transitive dependents |
| Value representation | String boxing | `Value` tagged union, no boxing for numbers/bools |
| Operator evaluation | String-keyed map lookup per operator | Switch on opcode byte |
| Stack implementation | Heap-allocated nodes | `[]Value` slice |

### Key Performance Properties

1. **Compiled formulas**: No re-tokenizing or re-parsing. The bytecode is a compact `[]Instruction` walked in a tight loop, cached on the `Cell`.
2. **No reflection**: Function dispatch via switch on function ID.
3. **Incremental recalculation**: Changing one cell only marks the affected subgraph dirty via BFS on the dependency graph, not all formulas.
4. **Lazy evaluation**: Dirty formula cells are only re-evaluated when their value is read (via `GetValue`), or eagerly via `Recalculate()`.
5. **Generation counters**: The `calcGen`/`cachedGen` mechanism provides a cheap staleness check without requiring every formula cell to be visited on mutation.

---

## Implementation Phases

### Phase 1: Core Read/Write ✅
- [x] OOXML reader/writer (workbook, worksheets, shared strings, styles)
- [x] Cell value get/set (numbers, strings, bools, dates)
- [x] Sheet operations (create, delete, list)
- [x] Row/cell iteration (Go 1.23 `iter.Seq` iterators)
- [ ] Basic streaming reader/writer

#### Test Coverage (root package): 77.2%

Tests cover: round-trip read/write, shared string deduplication, sparse data, multi-sheet operations, sheet deletion, duplicate sheet names, formula evaluation and round-trip, formula caching and dirty detection, dependency graph integration, row/cell iteration, MaxRow/MaxCol, SetValue/GetValue, date values, ZIP validity. The `ooxml/` internal package currently has 0% direct test coverage (exercised indirectly through integration tests).

### Phase 2: Formula Engine ✅
- [x] Lexer (`formula/lexer.go`) — tokenizer with 32 tests
- [x] Parser (`formula/parser.go`) — Pratt parser producing AST, with comprehensive tests
  - [x] AST node types (`formula/ast.go`): `NumberLit`, `StringLit`, `BoolLit`, `ErrorLit`, `CellRef`, `RangeRef`, `UnaryExpr`, `BinaryExpr`, `PostfixExpr`, `FuncCall`, `ArrayLit`
  - [x] `ErrorCode` type and constants (`formula/ast.go`)
  - [x] Cell reference parsing (`formula/cellref.go`): bare, absolute, mixed, sheet-qualified, quoted sheets with escape handling
  - [x] Operator precedence matching Excel: `^` right-associative, unary `-`/`+` so `-2^3 = -(2^3)`, greedy postfix `%`
  - [x] S-expression `String()` on all nodes for debugging and test output
- [x] Compiler (`formula/compiler.go`) — AST to bytecode with constant/ref deduplication, 99 registered functions
- [x] VM evaluator (`formula/eval.go`) — stack-based VM, `CellResolver` interface for cell/range lookup
- [x] Core function set (99 functions implemented across 7 files):
  - [x] Math (27): ABS, ACOS, ASIN, ATAN, ATAN2, CEILING, COS, EXP, FLOOR, INT, LN, LOG, LOG10, MOD, PI, POWER, PRODUCT, RAND, RANDBETWEEN, ROUND, ROUNDDOWN, ROUNDUP, SIN, SQRT, TAN
  - [x] Statistics (16): SUM, AVERAGE, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, LARGE, MAX, MEDIAN, MIN, SMALL, SUMIF, SUMIFS, SUMPRODUCT
  - [x] Text (22): CHAR, CHOOSE, CLEAN, CODE, CONCATENATE/CONCAT, EXACT, FIND, LEFT, LEN, LOWER, MID, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, TEXT, TRIM, UPPER, VALUE
  - [x] Logic (7): IF, IFERROR, AND, OR, NOT, XOR, SORT
  - [x] Lookup (7): VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP, INDIRECT (stub)
  - [x] Info (11): ISBLANK, ISERR, ISERROR, ISNA, ISNUMBER, ISTEXT, IFNA, COLUMN, ROW, COLUMNS, ROWS
  - [x] Date (11): DATE, DAY, HOUR, MINUTE, MONTH, NOW, SECOND, TIME, TODAY, YEAR
- [x] Cell reference resolution (single cell, ranges, cross-sheet via `CellResolver` interface)
- [x] Formula caching with generation counters (`calcGen`/`cachedGen`) for lazy evaluation
- [x] Dependency graph (`formula/depgraph.go`) with BFS transitive invalidation
- [x] Incremental recalculation — `SetValue` marks only affected dependents dirty
- [x] Circular reference detection (via `evaluating` map at eval time)
- [x] `Recalculate()` for eager evaluation of all dirty cells
- [ ] Topological sort for optimal recalculation order (Kahn's algorithm)
- [ ] Iterative circular reference convergence (MaxIterations/IterationTolerance)

#### Test Coverage (formula package): 82.9%

The formula engine has comprehensive test coverage across 13 test files:

- **Lexer tests** (`lexer_test.go`): 32+ tests covering all token types, edge cases
- **Parser tests** (`parser_test.go`): AST construction for all node types
- **Cell ref tests** (`cellref_test.go`): absolute, relative, mixed, sheet-qualified, quoted sheet names
- **Compiler tests** (`compiler_test.go`): bytecode generation for literals, refs, ranges, operators, functions, arrays, constant deduplication
- **Eval tests** (`eval_test.go`): arithmetic, cell references, ranges, string concat, comparisons, division edge cases, type coercion (`coerceNum` with empty/string/bool/error), `compareValues` (same-type and cross-type ordering), `isTruthy` (all value types), `valueToString`/`errorValueToString` (all error codes), error propagation through all operators, unary/percent edge cases, large number arithmetic, empty cell handling, array literals, ROW/COLUMN/ROWS/COLUMNS, IFNA
- **Dep graph tests** (`depgraph_test.go`): register/unregister, direct dependents, transitive BFS, range containment
- **Math function tests** (`functions_math_test.go`): basic correctness for trig, rounding, log, power functions
- **Stat function tests** (`functions_stat_test.go`): MEDIAN, LARGE/SMALL (including k out of range), COUNTBLANK, SUMIF/COUNTIF (with operator and wildcard criteria), SUMPRODUCT, COUNTA, SUMIFS/COUNTIFS/AVERAGEIF/AVERAGEIFS (multiple criteria, no-match DIV/0), SUM with mixed types, MIN/MAX with negatives/empty, error propagation in ranges, extended `matchesCriteria` (wildcards, case-insensitive, numeric operators)
- **Text function tests** (`functions_text_test.go`): CHAR/CODE (bounds), CLEAN, PROPER, REPLACE (insert/delete/boundary), REPT (zero/negative), FIND/SEARCH (case sensitivity, start position, not found), LEFT/RIGHT/MID (default args, zero, exceeds length, negative), SUBSTITUTE (all occurrences, specific instance, no match, case sensitive), TEXT format codes (decimals, percent, commas, negative), VALUE (commas, dollar, percent, whitespace, non-numeric), EXACT (case-sensitive), CHOOSE (out of range), CONCATENATE with mixed types, LEN with Unicode
- **Lookup function tests** (`functions_lookup_test.go`): VLOOKUP (exact, approximate, not found, col index out of range, string keys, case-insensitive), HLOOKUP (approximate, not found, row index out of range), MATCH (exact/ascending/descending, not found), INDEX (row/col out of range, 2-arg form), INDEX+MATCH combo, LOOKUP (3-arg vector form, approximate, too small), XLOOKUP (exact, next smaller, next larger, if_not_found, default #N/A)
- **Logic function tests** (`functions_logic_test.go`): AND/OR/NOT, XOR (odd/even true count), IF (2-arg missing else)
- **Info function tests** (`functions_info_test.go`): ISBLANK/ISNUMBER/ISTEXT, ISERROR (div/0, #N/A, #VALUE!, non-error)
- **Date function tests** (`functions_date_test.go`): DATE, DAY, MONTH, YEAR, TODAY, NOW

Root-level integration tests for the formula engine:
- **Formula eval tests** (`formula_eval_test.go`): end-to-end formula evaluation via `Sheet.GetValue`
- **Formula cache tests** (`formula_cache_test.go`): generation counter caching, dirty/stale detection, `Recalculate()`
- **Formula round-trip tests** (`formula_roundtrip_test.go`): formula text and cached values survive XLSX serialization

### Phase 3: Extended Functions & Features
- [ ] Remaining ~450 formula functions
- [ ] Array formulas and dynamic arrays
- [ ] Defined names / named ranges
- [ ] Data validation
- [ ] Conditional formatting
- [ ] Merged cells
- [ ] Auto-filters and tables

### Phase 4: Advanced Features
- [ ] Charts (read/write, not render)
- [ ] Pivot tables
- [ ] Sparklines
- [ ] Comments / notes
- [ ] Images
- [ ] File encryption/decryption

---

## Constraints & Limits

Match Excel limits:
- Max rows: 1,048,576
- Max columns: 16,384
- Max characters per cell: 32,767
- Max sheet name length: 31 characters
- Max formula length: 8,192 characters
