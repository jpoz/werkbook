# Werkbook Architecture Specification

## Overview

Werkbook is a Go library for reading, writing, and recalculating Excel XLSX files. It prioritizes fast recalculation through a compiled formula engine with a proper dependency graph, while maintaining full compatibility with the Office Open XML (OOXML) standard.

### Design Principles

1. **Fast recalculation** - Formulas are parsed once, compiled to bytecode, and evaluated via a stack-based VM. A dependency graph enables incremental recalculation.
2. **Correct by default** - Match Excel behavior for formula evaluation, type coercion, and edge cases.
3. **Memory efficient** - Stream large files; lazy-load sheets; spill to disk when needed.
4. **Simple API** - Provide a clean, idiomatic Go API that covers the common 90% of use cases without exposing OOXML internals.

---

## Package Structure

```
werkbook/
  werkbook.go          # File struct, Open, New, Save
  sheet.go             # Sheet type and operations
  cell.go              # Cell type, getting/setting values
  row.go               # Row iteration and access
  style.go             # Style definitions and application
  formula/
    lexer.go           # Formula tokenizer
    parser.go          # Token stream -> AST
    compiler.go        # AST -> bytecode
    vm.go              # Bytecode evaluator (the calc VM)
    functions.go       # Built-in function registry
    functions_math.go  # Math function implementations
    functions_text.go  # Text function implementations
    functions_stat.go  # Statistical function implementations
    functions_date.go  # Date/time function implementations
    functions_lookup.go # Lookup function implementations
    functions_logic.go # Logical function implementations
    functions_fin.go   # Financial function implementations
    functions_eng.go   # Engineering function implementations
    functions_info.go  # Information function implementations
    depgraph.go        # Cell dependency graph
    reference.go       # Cell/range reference parsing and resolution
    types.go           # Value types (Number, String, Bool, Error, Array)
  ooxml/
    reader.go          # ZIP/XML reading
    writer.go          # ZIP/XML writing
    workbook.go        # xlsxWorkbook structs
    worksheet.go       # xlsxWorksheet structs
    styles.go          # xlsxStyleSheet structs
    sharedstrings.go   # Shared string table
    relationships.go   # Package relationships
    contenttypes.go    # Content type mappings
    calcchain.go       # Calculation chain XML (for Excel compat)
  stream/
    writer.go          # Streaming row writer for large files
    reader.go          # Streaming row reader for large files
```

---

## Core Types

### File

The root object representing an open workbook.

```go
type File struct {
    sheets    []*Sheet
    sheetMap  map[string]*Sheet       // name -> sheet
    styles    *StyleSheet
    strings   *SharedStrings
    depGraph  *formula.DepGraph       // global dependency graph
    options   Options
}

type Options struct {
    // Memory management
    MaxMemoryMB       int  // spill to disk above this (default 256)

    // Calculation
    MaxIterations     int  // for circular refs (default 100)
    IterationTolerance float64 // convergence threshold (default 0.001)
}
```

### Sheet

```go
type Sheet struct {
    file     *File
    name     string
    index    int
    rows     map[int]*Row   // sparse row storage, keyed by 1-based row number
    cols     []ColDef       // column width/style defaults
    merges   []CellRange
    maxRow   int
    maxCol   int
}
```

### Row and Cell

```go
type Row struct {
    sheet  *Sheet
    num    int             // 1-based row number
    cells  map[int]*Cell   // sparse cell storage, keyed by 1-based column
    height float64
    hidden bool
}

type Cell struct {
    row      *Row
    col      int            // 1-based column number

    // Exactly one of these is set:
    value    Value           // literal value (number, string, bool, error)
    formula  *CompiledFormula // non-nil if this cell has a formula

    style    int             // index into StyleSheet
    cached   Value           // last calculated value (for formula cells)
    dirty    bool            // needs recalculation
}
```

### Value

A tagged union for cell values. Used as operands throughout the formula engine.

```go
type Value struct {
    Type    ValueType
    Num     float64
    Str     string
    Bool    bool
    Err     ErrorCode
    Array   [][]Value  // for array formulas / ranges
}

type ValueType byte

const (
    TypeEmpty ValueType = iota
    TypeNumber
    TypeString
    TypeBool
    TypeError
    TypeArray
)

type ErrorCode byte

const (
    ErrNone ErrorCode = iota
    ErrDIV0     // #DIV/0!
    ErrNA       // #N/A
    ErrNAME     // #NAME?
    ErrNULL     // #NULL!
    ErrNUM      // #NUM!
    ErrREF      // #REF!
    ErrVALUE    // #VALUE!
    ErrCALC     // #CALC!
)
```

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

### Phase 4: VM (`formula/vm.go`)

A stack-based virtual machine that evaluates compiled formulas. The hot loop is a switch over opcodes with no reflection.

```go
type VM struct {
    file     *File
    stack    []Value    // operand stack, pre-allocated
    sp       int        // stack pointer
    ctx      *CalcContext
}

type CalcContext struct {
    entry      CellAddr
    iterations map[CellAddr]int
    cache      map[CellAddr]Value
}

func (vm *VM) Eval(f *CompiledFormula, sheet *Sheet, row, col int) (Value, error) {
    vm.sp = 0
    for _, instr := range f.Code {
        switch instr.Op {
        case OpPushNum:
            vm.push(f.Consts[instr.Operand])
        case OpLoadCell:
            addr := f.Refs[instr.Operand]
            vm.push(vm.resolveCell(addr))
        case OpLoadRange:
            rng := f.Ranges[instr.Operand]
            vm.push(vm.resolveRange(rng))
        case OpAdd:
            b, a := vm.pop(), vm.pop()
            vm.push(addValues(a, b))
        case OpCall:
            funcID := instr.Operand >> 8
            argc := instr.Operand & 0xFF
            args := vm.popN(int(argc))
            vm.push(builtinFuncs[funcID](args))
        // ... other opcodes
        }
    }
    return vm.stack[0], nil
}
```

Key performance properties:
- **No allocations in the hot path** - the stack is pre-allocated, Values are passed by value
- **No reflection** - function dispatch via a function pointer table indexed by function ID
- **No re-parsing** - bytecode is compiled once per formula and reused across recalculations
- **Cache-friendly** - instructions are a compact byte stream; the operand stack is a contiguous slice

### Function Registry (`formula/functions.go`)

Functions are registered in a table at init time. No reflection.

```go
type BuiltinFunc func(args []Value) Value

var builtinFuncs []BuiltinFunc   // indexed by function ID
var funcNameToID map[string]int  // "SUM" -> 0, "IF" -> 1, ...

func init() {
    register("SUM", fnSUM)
    register("IF", fnIF)
    register("VLOOKUP", fnVLOOKUP)
    // ... 500+ functions
}

func register(name string, fn BuiltinFunc) {
    id := len(builtinFuncs)
    builtinFuncs = append(builtinFuncs, fn)
    funcNameToID[name] = id
}
```

---

## Dependency Graph (`formula/depgraph.go`)

This is the core of fast recalculation. Werkbook tracks which cells depend on which other cells and only recalculates what is necessary, rather than clearing all caches on any mutation.

### Data Structure

```go
type CellAddr struct {
    Sheet int // sheet index
    Col   int // 1-based
    Row   int // 1-based
}

type DepGraph struct {
    mu         sync.RWMutex

    // Forward edges: "cell X depends on cells Y, Z, ..."
    // Used to invalidate: when Y changes, look up who depends on Y.
    dependents map[CellAddr][]CellAddr  // Y -> [cells that reference Y]

    // Reverse edges: "cell X references cells Y, Z, ..."
    // Used during formula registration/removal.
    references map[CellAddr][]CellAddr  // X -> [cells that X references]

    // Range dependents: cells that depend on a range (e.g., SUM(A1:A100))
    // Stored separately because a change to any cell in the range must
    // trigger recalculation of the dependent.
    rangeDeps  map[CellAddr][]RangeAddr
    rangeSubs  map[rangeKey][]CellAddr   // range -> [cells that reference this range]
}
```

### Operations

**Register a formula:**
When `SetCellFormula` is called, the formula is compiled, and `Refs`/`Ranges` from the compiled formula are used to register edges in the dependency graph.

```go
func (dg *DepGraph) Register(cell CellAddr, compiled *CompiledFormula) {
    // Remove old edges for this cell
    dg.Unregister(cell)

    // Add forward edges for cell refs
    for _, ref := range compiled.Refs {
        dg.dependents[ref] = append(dg.dependents[ref], cell)
    }
    // Add range subscriptions
    for _, rng := range compiled.Ranges {
        key := rangeKey{rng.Sheet, rng.FromCol, rng.FromRow, rng.ToCol, rng.ToRow}
        dg.rangeSubs[key] = append(dg.rangeSubs[key], cell)
    }
    dg.references[cell] = compiled.Refs
    dg.rangeDeps[cell] = compiled.Ranges
}
```

**Invalidate on mutation:**
When a cell value changes, walk the dependency graph to mark all transitive dependents as dirty.

```go
func (dg *DepGraph) Invalidate(changed CellAddr) []CellAddr {
    var dirty []CellAddr
    visited := make(map[CellAddr]bool)
    queue := []CellAddr{changed}

    for len(queue) > 0 {
        cur := queue[0]
        queue = queue[1:]
        if visited[cur] { continue }
        visited[cur] = true

        // Direct cell dependents
        for _, dep := range dg.dependents[cur] {
            dirty = append(dirty, dep)
            queue = append(queue, dep)
        }

        // Range dependents: any range that contains `cur`
        for key, deps := range dg.rangeSubs {
            if key.Contains(cur) {
                for _, dep := range deps {
                    if !visited[dep] {
                        dirty = append(dirty, dep)
                        queue = append(queue, dep)
                    }
                }
            }
        }
    }
    return dirty
}
```

**Topological sort for recalculation order:**
After collecting dirty cells, sort them so that dependencies are evaluated before dependents.

```go
func (dg *DepGraph) CalcOrder(dirty []CellAddr) ([]CellAddr, error) {
    // Kahn's algorithm for topological sort
    // Returns ErrCircularRef if a cycle is detected (handled via iteration)
}
```

### Recalculation Flow

```
1. User calls SetCellValue(sheet, "A1", 42)
2.   -> cell.value = 42, cell.cached = 42
3.   -> depGraph.Invalidate(A1) returns [B1, C1, D5] (cells with formulas referencing A1)
4.   -> depGraph.CalcOrder([B1, C1, D5]) returns [B1, C1, D5] (topo sorted)
5.   -> for each cell in order:
6.        cell.dirty = true
7. User calls CalcCellValue("Sheet1", "D5") or Recalculate()
8.   -> if cell.dirty:
9.        vm.Eval(cell.formula, ...) -> new value
10.       cell.cached = new value
11.       cell.dirty = false
```

This means:
- Changing a cell value does NOT trigger recalculation (just marks dirty cells).
- Reading a formula cell's value triggers lazy evaluation if dirty.
- `Recalculate()` eagerly evaluates all dirty cells in dependency order.
- Only the affected subgraph is recalculated, not the entire workbook.

### Circular Reference Handling

Circular references are detected during topological sort. When detected, the involved cells are evaluated iteratively up to `Options.MaxIterations`, converging when successive values differ by less than `Options.IterationTolerance`.

---

## Reading XLSX Files (`ooxml/reader.go`)

### Process

1. Open ZIP archive, enumerate entries
2. Parse `[Content_Types].xml` to discover part types
3. Parse `_rels/.rels` for package relationships
4. Parse `xl/workbook.xml` -> workbook metadata, sheet list
5. For each sheet, parse `xl/worksheets/sheet{N}.xml` -> rows and cells
6. Parse `xl/sharedStrings.xml` -> string table (lazy, on first string access)
7. Parse `xl/styles.xml` -> style definitions
8. Parse `xl/calcChain.xml` -> used only for initial dirty-marking

### Lazy Sheet Loading

Sheets are not parsed until first access. The ZIP entry's byte content is held in memory (or temp file for large entries), and deserialization happens on `file.Sheet("Sheet1")`.

### Large File Support

- ZIP entries larger than `Options.MaxMemoryMB / sheetCount` are extracted to temp files
- Streaming reader (`stream/reader.go`) provides a row iterator that parses XML incrementally without loading the full sheet

---

## Writing XLSX Files (`ooxml/writer.go`)

### Process

1. Create ZIP writer
2. Serialize workbook.xml from `File.sheets` metadata
3. For each sheet, serialize worksheet XML from `Sheet.rows`
4. Serialize shared strings (deduplicated during write)
5. Serialize styles
6. Serialize calculation chain (from depGraph for Excel compat)
7. Write content types and relationships
8. Close ZIP

### Streaming Writer (`stream/writer.go`)

For writing large files without holding the full sheet in memory:

```go
sw, _ := file.NewStreamWriter("Sheet1")
for i := 1; i <= 1_000_000; i++ {
    sw.SetRow(i, []any{i, "hello", 3.14})
}
sw.Flush()
file.SaveAs("big.xlsx")
```

Rows are serialized to XML immediately and buffered to a temp file. On save, the temp file content is streamed directly into the ZIP entry.

---

## API Surface

### Opening and Creating

```go
// Open an existing file
f, err := werkbook.Open("report.xlsx")
f, err := werkbook.Open("report.xlsx", werkbook.Options{MaxMemoryMB: 512})

// Create a new file
f := werkbook.New()

// Save
err := f.SaveAs("output.xlsx")
err := f.Save() // overwrite original
```

### Sheet Operations

```go
sheet := f.Sheet("Sheet1")
sheet, err := f.NewSheet("Data")
f.DeleteSheet("Sheet2")
names := f.SheetNames()
```

### Cell Access

```go
// Get/set values
val, err := sheet.Cell("A1")           // returns Value
sheet.SetValue("A1", 42)               // sets number
sheet.SetValue("A1", "hello")          // sets string
sheet.SetValue("A1", true)             // sets bool
sheet.SetValue("A1", time.Now())       // sets date

// Formulas
sheet.SetFormula("B1", "SUM(A1:A100)")
formula, err := sheet.Formula("B1")    // returns "SUM(A1:A100)"

// Calculated value (lazy - evaluates if dirty)
result, err := sheet.CalcValue("B1")   // returns Value

// Bulk read
for row := range sheet.Rows() {        // iterator
    for _, cell := range row.Cells() {
        // ...
    }
}
```

### Recalculation

```go
// Recalculate all dirty cells
err := f.Recalculate()

// Recalculate a single cell (and its dependencies)
val, err := f.CalcCell("Sheet1", "B1")
```

### Styles

```go
style := werkbook.Style{
    Font:      &werkbook.Font{Bold: true, Size: 12, Name: "Calibri"},
    Fill:      &werkbook.Fill{Type: "solid", Color: "#FF0000"},
    Alignment: &werkbook.Alignment{Horizontal: "center"},
    NumFmt:    "#,##0.00",
}
id, err := f.NewStyle(style)
sheet.SetStyle("A1:D1", id)
```

---

## Performance Design

| Operation | Naive Approach | Werkbook |
|-----------|----------------|----------|
| Parse formula | Every evaluation (re-tokenize) | Once (compile to bytecode, cache) |
| Function dispatch | Reflection-based | Direct function pointer lookup by index |
| Dependency tracking | None (clear all caches on any mutation) | DAG with incremental invalidation |
| Recalc scope | All formula cells | Only dirty cells + transitive dependents |
| Value representation | String boxing | `Value` tagged union, no boxing for numbers/bools |
| Operator evaluation | String-keyed map lookup per operator | Switch on opcode byte |
| Range resolution | Materialize full matrix | Lazy resolution, stream from sparse row map |
| Defined name lookup | Linear scan of all names | Hash map lookup |
| Stack implementation | Heap-allocated nodes | Pre-allocated `[]Value` slice |

### Expected Speedup Sources

1. **Compiled formulas** (~5-10x for repeated evaluation): No re-tokenizing or re-parsing. The bytecode is a compact `[]Instruction` walked in a tight loop.
2. **No reflection** (~2-3x for function-heavy sheets): Direct indexed function calls instead of reflection-based dispatch.
3. **Incremental recalculation** (unbounded, depends on workbook): Changing one cell in a million-row sheet only recalculates the affected subgraph, not all formulas.
4. **Cache-friendly data structures** (~1.5-2x): Pre-allocated operand stack, compact instruction encoding, sparse row/cell maps.
5. **Lazy range evaluation** (~2-5x for large ranges): Avoid materializing `A:A` as a 1M-element matrix when SUM just needs to iterate.

---

## Implementation Phases

### Phase 1: Core Read/Write ✅
- [x] OOXML reader/writer (workbook, worksheets, shared strings, styles)
- [x] Cell value get/set (numbers, strings, bools, dates)
- [x] Sheet operations (create, delete, rename, list)
- [x] Row/column operations
- [ ] Basic streaming reader/writer

### Phase 2: Formula Engine (in progress)
- [x] Lexer (`formula/lexer.go`) — tokenizer with 32 tests
- [x] Parser (`formula/parser.go`) — Pratt parser producing AST, with comprehensive tests
  - [x] AST node types (`formula/ast.go`): `NumberLit`, `StringLit`, `BoolLit`, `ErrorLit`, `CellRef`, `RangeRef`, `UnaryExpr`, `BinaryExpr`, `PostfixExpr`, `FuncCall`, `ArrayLit`
  - [x] `ErrorCode` type and constants (`formula/ast.go`)
  - [x] Cell reference parsing (`formula/cellref.go`): bare, absolute, mixed, sheet-qualified, quoted sheets with escape handling
  - [x] Operator precedence matching Excel: `^` right-associative, unary `-`/`+` so `-2^3 = -(2^3)`, greedy postfix `%`
  - [x] S-expression `String()` on all nodes for debugging and test output
- [ ] Compiler (AST -> bytecode)
- [ ] VM evaluator
- [ ] Core function set (~50 most common: SUM, IF, VLOOKUP, INDEX, MATCH, AVERAGE, COUNT, MIN, MAX, CONCATENATE, LEFT, RIGHT, MID, DATE, TODAY, NOW, AND, OR, NOT, IFERROR, ROUND, ABS, SUMIF, COUNTIF, SUMPRODUCT, TRIM, LEN, SUBSTITUTE, TEXT, VALUE, LARGE, SMALL, RANK, OFFSET, INDIRECT)
- [ ] Cell reference resolution (single cell, ranges, cross-sheet)
- [ ] Dependency graph and incremental invalidation
- [ ] Circular reference detection and iterative evaluation

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
