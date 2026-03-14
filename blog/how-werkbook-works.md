# How Werkbook Works: A Deep Dive into a Pure Go Spreadsheet Engine

Spreadsheets are everywhere. They power business operations, financial models, scientific data analysis, and countless other workflows. Yet when developers need to work with `.xlsx` files programmatically, they often reach for libraries that are either incomplete, bloated with dependencies, or lack one critical feature: the ability to actually *evaluate formulas*.

Werkbook takes a different approach. It's a pure Go library — zero external dependencies — that can read, write, and *calculate* XLSX spreadsheets. This post takes you deep inside its architecture, from the two-layer design that cleanly separates concerns, to the bytecode virtual machine that powers its formula engine, to the dependency graph that makes incremental recalculation possible.

---

## The Big Picture: Two Layers, One Goal

Werkbook's architecture is built around a clean separation of concerns. At the highest level, there are two layers:

1. **The Public API Layer** — the types you interact with as a developer: `File`, `Sheet`, `Row`, `Cell`, `Value`. These use 1-based indexing (Excel-style) and provide an intuitive interface for working with spreadsheet data.

2. **The OOXML Layer** — the serialization engine that translates between the in-memory model and the actual `.xlsx` file format (which is really a ZIP archive containing XML documents).

Between these two layers sits the **Formula Engine**, a self-contained subsystem with its own lexer, parser, compiler, and virtual machine. It's the heart of what makes werkbook more than just an XML shuffler.

```
┌─────────────────────────────────────────────┐
│            Public API Layer                 │
│   File · Sheet · Row · Cell · Value         │
├─────────────────────────────────────────────┤
│            Formula Engine                   │
│   Lexer → Parser → Compiler → VM           │
│   Dependency Graph · Function Registry      │
├─────────────────────────────────────────────┤
│            OOXML Layer                      │
│   ZIP/XML · SharedStrings · Styles          │
└─────────────────────────────────────────────┘
```

Let's work through each of these in detail.

---

## The In-Memory Model

When you open or create a workbook in werkbook, you're working with a tree of Go structs that represent the spreadsheet in memory.

### File: The Workbook Container

The in-memory object tree looks like this:

```
                         ┌──────────┐
                         │   File   │
                         └────┬─────┘
               ┌──────────────┼──────────────┐
               │              │              │
          ┌────▼────┐   ┌─────▼─────┐  ┌─────▼─────┐
          │ Sheet 1 │   │  Sheet 2  │  │  Sheet N  │
          └────┬────┘   └───────────┘  └───────────┘
               │
      ┌────────┼────────┐
      │        │        │
  ┌───▼──┐ ┌──▼───┐ ┌──▼───┐
  │ Row 1│ │ Row 2│ │ Row N│
  └───┬──┘ └──┬───┘ └──────┘
      │        │
  ┌───▼──┐ ┌──▼───┐
  │Cell  │ │Cell  │  ...
  │A1    │ │A2    │
  └──────┘ └──────┘
```

The `File` struct is the root of everything:

```go
type File struct {
    sheets         []*Sheet
    sheetNames     []string
    date1904       bool
    calcProps      CalcProperties
    coreProps      CoreProperties
    calcGen        uint64
    evaluating     map[cellKey]bool
    deps           *formula.DepGraph
    tableDefs      []Table
    tables         []formula.TableInfo
    definedNames   []formula.DefinedNameInfo
}
```

A few things stand out here:

- **`calcGen`** is a generation counter that starts at 1 and increments every time any cell is mutated. This is the backbone of lazy recalculation — a cell's cached value is valid only if its `cachedGen` matches the file's `calcGen`.

- **`evaluating`** is a map used for circular reference detection. When a formula is being evaluated, its cell key is added to this map. If the evaluator encounters that same key again during resolution, it knows it has hit a cycle and returns a `#REF!` error.

- **`deps`** is the dependency graph that tracks which formula cells depend on which data cells, enabling incremental recalculation.

### Sheet, Row, and Cell: Sparse by Design

Sheets use sparse data structures — maps rather than dense arrays:

```go
// Conceptually:
Sheet.rows = map[int]*Row    // row number → Row
Row.cells  = map[int]*Cell   // column number → Cell
```

This means a sheet with data in A1 and Z1000 doesn't allocate memory for the 25,998 empty cells between them. Compare dense vs. sparse storage:

```
Dense array (what werkbook does NOT do):
┌─────┬───┬───┬───┬───┬───┬─── ── ──┬───┐
│     │ A │ B │ C │ D │ E │ ...     │ Z │
├─────┼───┼───┼───┼───┼───┼─── ── ──┼───┤
│   1 │ x │   │   │   │   │         │   │  ← 26 cells allocated per row
│   2 │   │   │   │   │   │         │   │  ← even if only 1 has data
│   3 │   │   │   │   │   │         │   │
│ ... │   │   │   │   │   │         │   │  25,974 empty cells wasting memory
│1000 │   │   │   │   │   │         │ x │
└─────┴───┴───┴───┴───┴───┴─── ── ──┴───┘

Sparse map (what werkbook DOES):
rows: {
  1 → cells: { 1 → Cell{A1} }       ← only 1 entry
  1000 → cells: { 26 → Cell{Z1000} } ← only 1 entry
}
Total: 2 rows, 2 cells allocated  ✓
```

It's a natural fit for the way real spreadsheets are used, where data tends to cluster in specific regions.

Each `Cell` holds its value, its formula text (if any), a compiled formula (lazily computed), the generation at which its cached value was computed, a dirty flag for dependency-based invalidation, and an optional style:

```go
type Cell struct {
    col            int
    value          Value
    formula        string
    isArrayFormula bool
    formulaRef     string
    compiled       *formula.CompiledFormula
    cachedGen      uint64
    dirty          bool
    style          *Style
}
```

### The Value Type: A Tagged Union

Werkbook's `Value` type is a tagged union that can hold any of the types a spreadsheet cell can contain — numbers, strings, booleans, errors, dates, and the empty value. Rather than using Go interfaces (which would require heap allocation for every cell value), it uses a struct with a type tag:

This approach avoids boxing primitive types into interfaces, which keeps memory allocation tight — an important consideration when a workbook can contain millions of cells.

---

## Reading XLSX Files: Unzipping the Onion

An `.xlsx` file is a ZIP archive containing a specific structure of XML documents. When you call `werkbook.Open("file.xlsx")`, the OOXML layer handles the unglamorous but essential work of parsing this structure.

### The OOXML File Structure

```
file.xlsx (ZIP archive)
├── [Content_Types].xml          # MIME type declarations
├── _rels/.rels                  # Root relationships
├── docProps/
│   └── core.xml                 # Author, title, dates
└── xl/
    ├── workbook.xml             # Sheet names and ordering
    ├── sharedStrings.xml        # Deduplicated string table
    ├── styles.xml               # Fonts, fills, borders, number formats
    ├── _rels/workbook.xml.rels  # Sheet file references
    └── worksheets/
        ├── sheet1.xml           # Cell data for Sheet1
        ├── sheet2.xml           # Cell data for Sheet2
        └── ...
```

### The Shared String Table

One of XLSX's optimizations is the **Shared String Table (SST)**. Instead of repeating the string "Revenue" in every cell that contains it, the file stores it once in `sharedStrings.xml` and references it by index:

```
sharedStrings.xml              worksheet1.xml
┌─────────────────┐            ┌──────────────────────────┐
│ 0: "Revenue"    │◄───────────│ A1: type="s" value="0"   │
│ 1: "Expenses"   │◄───────────│ A2: type="s" value="1"   │
│ 2: "Profit"     │◄───────────│ A3: type="s" value="2"   │
└─────────────────┘       ┌────│ B1: type="s" value="0"   │
                          │    └──────────────────────────┘
                          │
                          └──── Same string "Revenue" —
                               stored once, referenced twice
```

During reading, werkbook resolves these indices back to actual strings. During writing, it builds a new SST by deduplicating all string values across the workbook.

### The Reading Pipeline

```
werkbook.Open("file.xlsx")
  │
  ├── ooxml.ReadWorkbook()
  │     ├── Open ZIP archive
  │     ├── Parse xl/workbook.xml → sheet names/order
  │     ├── Parse xl/_rels/workbook.xml.rels → sheet file paths
  │     ├── Parse xl/sharedStrings.xml → string lookup table
  │     ├── Parse xl/styles.xml → fonts, fills, borders, formats
  │     ├── For each sheet:
  │     │     Parse xl/worksheets/sheetN.xml → cell data
  │     └── Parse docProps/core.xml → metadata
  │     └── Return WorkbookData (intermediate representation)
  │
  └── fileFromData(WorkbookData)
        ├── Create File, Sheet, Row, Cell objects
        ├── Resolve SST indices → actual string values
        ├── Mark formula cells with cached generation
        ├── Register formula dependencies in DepGraph
        └── Return ready-to-use *File
```

The intermediate `WorkbookData` type is the bridge between the XML world and the API world. It exists so that the OOXML package never leaks XML-specific details into the public API, and the public API never needs to know about XML namespaces or ZIP entry paths.

---

## Writing XLSX Files: Rebuilding the Archive

When you call `File.SaveAs("output.xlsx")`, the process runs in reverse — but with a few important wrinkles.

### Formula Recalculation Before Save

Before serializing, werkbook recalculates any dirty formulas. This ensures the saved file contains up-to-date cached values, which is important because some spreadsheet applications (particularly lightweight viewers) don't have their own formula engines and rely on cached values.

### Style Deduplication

XLSX stores styles in a shared pool. If 1,000 cells have bold text, there's one font definition and 1,000 cells referencing it by index:

```
styles.xml                         worksheet.xml
┌────────────────────────┐         ┌─────────────────────┐
│ fonts:                 │         │ A1: styleIdx=0      │──┐
│   0: {bold, 12pt}      │◄────────│ A2: styleIdx=0      │──┘
│   1: {italic, 10pt}    │◄────────│ B1: styleIdx=1      │
│                        │         │ B2: styleIdx=0      │──── same font
│ fills:                 │         │ C1: styleIdx=0      │──── reused, not
│   0: {yellow, solid}   │         │ ...                 │     duplicated
│   1: {none}            │         │ Z99: styleIdx=0     │
└────────────────────────┘         └─────────────────────┘
                                    1,000 cells → 2 font definitions
```

Werkbook handles this deduplication automatically during the write phase: it collects all unique fonts, fills, borders, alignments, and number formats, assigns each an index, and writes cells with references to those indices.

### The Writing Pipeline

```
File.SaveAs("output.xlsx")
  │
  ├── File.Recalculate()        # Ensure all formula values are current
  │
  ├── File.buildWorkbookData()
  │     ├── Deduplicate styles → index mapping
  │     ├── Convert dates to serial numbers
  │     ├── Build shared string table
  │     ├── Convert Cell values → CellData with type tags
  │     └── Return WorkbookData
  │
  └── ooxml.WriteWorkbook(WorkbookData)
        ├── Create ZIP writer
        ├── Write [Content_Types].xml
        ├── Write _rels/.rels
        ├── Write xl/workbook.xml
        ├── Write xl/worksheets/sheetN.xml for each sheet
        ├── Write xl/styles.xml (deduplicated)
        ├── Write xl/sharedStrings.xml
        ├── Write xl/_rels/workbook.xml.rels
        ├── Write docProps/core.xml
        └── Close ZIP → flush to file
```

---

## The Formula Engine: From Text to Bytecode to Results

The formula engine is the crown jewel of werkbook. It takes a formula string like `SUM(A1:A10)*1.08` and turns it into executable bytecode, evaluates it on a stack-based virtual machine, and returns a result — all while tracking dependencies for incremental recalculation.

The pipeline has four stages: **Lexing → Parsing → Compilation → Evaluation**.

```
                        The Formula Pipeline

  "SUM(A1:A10)*1.08"
         │
         ▼
  ┌─────────────┐     ┌─────┬─────┬───┬─────┬───┬──────┬───┬──────┐
  │    LEXER     │────▶│ SUM │  (  │A1 │  :  │A10│  )   │ * │ 1.08 │
  └─────────────┘     └─────┴─────┴───┴─────┴───┴──────┴───┴──────┘
                                       tokens
         │
         ▼
  ┌─────────────┐              ×
  │   PARSER    │────▶        / \
  │  (Pratt)    │          Call   1.08
  └─────────────┘          |
                          SUM
                           |              AST
                        Range
                        /    \
                      A1     A10
         │
         ▼
  ┌─────────────┐     ┌──────────────────────────────┐
  │  COMPILER   │────▶│ OpLoadRange  0               │
  └─────────────┘     │ OpCall       <SUM>           │
                      │ OpPushNum    0   (=1.08)     │
                      │ OpMul                        │  bytecode
                      └──────────────────────────────┘
         │
         ▼
  ┌─────────────┐
  │  VM (eval)  │────▶  Result: 11664.0
  │  stack-based│
  └─────────────┘
```

### Stage 1: Lexing

The lexer (`formula/lexer.go`) transforms a formula string into a stream of tokens. It handles all the quirks of Excel's formula syntax:

- **String literals:** `"Hello, ""World"""` (doubled quotes for escaping)
- **Numbers:** `3.14`, `.5`, `1E10`
- **Cell references:** `A1`, `$A$1`, `Sheet2!B5`
- **Range references:** `A1:C10`, `Sheet1:Sheet3!A1:B2`
- **Operators:** `+`, `-`, `*`, `/`, `^`, `&`, `=`, `<>`, `<`, `>`, `<=`, `>=`
- **Structured references:** `Table1[Column]`, `Table1[#Headers]`
- **Parentheses and commas** for function calls and grouping

The lexer is careful to distinguish between unary minus (negation) and binary minus (subtraction) based on context — a minus sign at the start of a formula or after an operator is unary.

Here's what the token stream looks like for a moderately complex formula:

```
Formula: IF(A1>0, A1*B1, "N/A")

Tokens:  ┌────┐┌───┐┌────┐┌───┐┌───┐┌───┐┌────┐┌───┐┌────┐┌───┐┌───────┐┌───┐
         │ IF ││ ( ││ A1 ││ > ││ 0 ││ , ││ A1 ││ * ││ B1 ││ , ││ "N/A" ││ ) │
         └────┘└───┘└────┘└───┘└───┘└───┘└────┘└───┘└────┘└───┘└───────┘└───┘
Type:    func  open  ref   op   num  sep  ref   op   ref   sep  str     close
```

### Stage 2: Parsing (Pratt Precedence Parser)

The parser (`formula/parser.go`) uses a **Pratt precedence parser** (also known as a top-down operator precedence parser) to build an Abstract Syntax Tree (AST). This approach is elegant because operator precedence and associativity are encoded as binding powers rather than grammar rules:

```
Precedence levels (lowest to highest):
  2: Comparison  (=, <>, <, >, <=, >=)
  4: Concatenation (&)
  6: Addition/Subtraction (+, -)
  8: Multiplication/Division (*, /)
 10: Exponentiation (^)
 14: Range (:) — highest precedence
```

The Pratt parser handles left and right associativity naturally. Exponentiation (`^`) is right-associative (so `2^3^4` means `2^(3^4)`), while arithmetic operators are left-associative (so `1-2-3` means `(1-2)-3`).

```
Left-associative: 1 - 2 - 3         Right-associative: 2 ^ 3 ^ 4

         -                                    ^
        / \                                  / \
       -   3                                2   ^
      / \                                      / \
     1   2                                    3   4

  = (1 - 2) - 3 = -4                 = 2 ^ (3 ^ 4) = 2⁸¹
```

The resulting AST contains node types for:
- **Literals:** numbers, strings, booleans, errors
- **Cell references:** single cells and ranges
- **Binary operations:** arithmetic, comparison, concatenation
- **Unary operations:** negation, percentage
- **Function calls:** with argument lists
- **Array literals:** `{1,2,3;4,5,6}`

Here's a more complex AST example for `IF(A1>0, A1*B1, "N/A")`:

```
              CallNode
             /   |   \
          "IF"  args:
               /  |  \
              /   |   \
          BinOp  BinOp  StringLit
          (>)    (*)     "N/A"
         / \    / \
       Ref  Num Ref  Ref
       A1   0   A1   B1
```

### Stage 3: Compilation (AST to Bytecode)

The compiler (`formula/compiler.go`) walks the AST and emits bytecode instructions. It uses **constant pooling** and **reference deduplication** to keep the bytecode compact:

```go
type CompiledFormula struct {
    Source string          // original formula text
    Code   []Instruction   // bytecode
    Consts []Value         // constant pool
    Refs   []CellAddr      // cell reference pool
    Ranges []RangeAddr     // range reference pool
}

type Instruction struct {
    Op      OpCode
    Operand uint32
}
```

If the same number appears multiple times in a formula, it's stored once in the constant pool and referenced by index. Same for cell references and ranges.

The full instruction set has **27 opcodes**:

| Category | Opcodes |
|----------|---------|
| Push | `OpPushNum`, `OpPushStr`, `OpPushBool`, `OpPushError`, `OpPushEmpty` |
| Load | `OpLoadCell`, `OpLoadRange`, `OpLoad3DRange`, `OpLoadCellRef` |
| Arithmetic | `OpAdd`, `OpSub`, `OpMul`, `OpDiv`, `OpPow`, `OpNeg`, `OpPercent` |
| Comparison | `OpEq`, `OpNe`, `OpLt`, `OpLe`, `OpGt`, `OpGe` |
| Other | `OpConcat`, `OpCall`, `OpMakeArray`, `OpEnterArrayCtx`, `OpLeaveArrayCtx`, `OpRefResultToBool` |

For example, the formula `SUM(A1:A10)*1.08` compiles to:

```
Bytecode:                     Constant Pool:       Range Pool:
┌────┬──────────────┬─────┐   ┌───┬────────┐       ┌───┬──────────┐
│ #  │ Opcode       │ Arg │   │ 0 │ 1.08   │       │ 0 │ A1:A10   │
├────┼──────────────┼─────┤   └───┴────────┘       └───┴──────────┘
│ 0  │ OpLoadRange  │  0  │─── loads range pool[0]
│ 1  │ OpCall       │ <S> │─── calls SUM (func ID)
│ 2  │ OpPushNum    │  0  │─── pushes const pool[0] = 1.08
│ 3  │ OpMul        │  -  │─── multiplies top two stack values
└────┴──────────────┴─────┘
```

### Stage 4: Evaluation (Stack-Based Virtual Machine)

The VM (`formula/eval.go`) is a classic stack machine. It processes instructions one at a time, pushing and popping values from a stack:

```go
func Eval(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext) (Value, error) {
    stack := make([]Value, 0, 16)
    // ... process each instruction
}
```

The `CellResolver` interface is how the VM accesses cell data without depending on the `Sheet` type directly:

```go
type CellResolver interface {
    GetCellValue(addr CellAddr) Value
    GetRangeValues(addr RangeAddr) [][]Value
}
```

This separation is key to testability — formula functions can be tested with mock resolvers without creating full workbook structures.

Let's visualize the VM executing `A1+A2*A3` where A1=2, A2=3, A3=4:

```
Bytecode: OpLoadCell 0(A1), OpLoadCell 1(A2), OpLoadCell 2(A3), OpMul, OpAdd

Step 1: OpLoadCell A1     Step 2: OpLoadCell A2     Step 3: OpLoadCell A3
┌─────┐                  ┌─────┐                  ┌─────┐
│     │                  │     │                  │  4  │ ← top
│     │                  │  3  │ ← top            ├─────┤
│  2  │ ← top            ├─────┤                  │  3  │
├─────┤                  │  2  │                  ├─────┤
│stack│                  ├─────┤                  │  2  │
└─────┘                  │stack│                  ├─────┤
                         └─────┘                  │stack│
                                                  └─────┘

Step 4: OpMul             Step 5: OpAdd
  pop 4 and 3,              pop 12 and 2,
  push 3*4=12               push 2+12=14
┌─────┐                  ┌─────┐
│     │                  │     │
│ 12  │ ← top            │ 14  │ ← top = result!
├─────┤                  ├─────┤
│  2  │                  │stack│
├─────┤                  └─────┘
│stack│
└─────┘
```

The VM also handles some subtle Excel behaviors:

- **Implicit intersection:** When a formula references an entire column (like `A:A`) in a non-array context, the VM intersects it with the current row, returning just the single value at that intersection point.

```
Implicit Intersection Example:

Cell C3 contains: =A:A + 1

        A:A (entire column)          After implicit intersection
       ┌─────┐                       at row 3:
  row 1│ 10  │
       ├─────┤                       A3 = 30
  row 2│ 20  │
       ├─────┤                       Result: 30 + 1 = 31
  row 3│ 30  │ ◄── current row
       ├─────┤
  row 4│ 40  │
       └─────┘
```

- **Type coercion:** Excel has complex implicit type conversion rules. In numeric contexts, the string "42" becomes the number 42, and `TRUE` becomes 1. In string contexts, the number 42 becomes "42". Werkbook faithfully reproduces these rules.

```
Type Coercion Rules:

  Numeric context (+ - * /)        String context (&)
  ┌───────────┬──────────┐         ┌───────────┬──────────┐
  │ Input     │ Becomes  │         │ Input     │ Becomes  │
  ├───────────┼──────────┤         ├───────────┼──────────┤
  │ "42"      │ 42       │         │ 42        │ "42"     │
  │ "3.14"    │ 3.14     │         │ TRUE      │ "TRUE"   │
  │ TRUE      │ 1        │         │ FALSE     │ "FALSE"  │
  │ FALSE     │ 0        │         │ #N/A      │ #N/A ✗   │
  │ "" (empty)│ 0        │         │ "" (empty)│ ""       │
  │ "hello"   │ #VALUE!  │         └───────────┴──────────┘
  └───────────┴──────────┘
```

- **Error propagation:** Most operations propagate errors — if one operand is `#DIV/0!`, the result is `#DIV/0!`. But some functions (like `IFERROR`) intentionally catch errors.

### The Function Registry: 438 Functions and Counting

Werkbook supports over 438 spreadsheet functions, organized into categories:

- **Math & Trigonometry:** `SUM`, `AVERAGE`, `ROUND`, `SIN`, `COS`, `LOG`, `MOD`, `RAND`, `CEILING`, `FLOOR`, and many more
- **Text:** `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `LEN`, `FIND`, `SUBSTITUTE`, `TRIM`, `UPPER`, `LOWER`, `TEXT`
- **Lookup & Reference:** `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH`, `OFFSET`, `INDIRECT`
- **Date & Time:** `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `EDATE`, `EOMONTH`, `NETWORKDAYS`
- **Logical:** `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`
- **Statistical:** `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, `MEDIAN`, `STDEV`
- **Financial:** `PMT`, `FV`, `PV`, `NPV`, `IRR`, `RATE`
- **Information:** `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISFORMULA`, `TYPE`
- **Engineering:** `BIN2DEC`, `DEC2BIN`, `HEX2DEC`, `COMPLEX`, `IMABS`
- **Array/Dynamic:** `SORT`, `SORTBY`, `FILTER`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`
- **Web:** `ENCODEURL`

Functions are registered at initialization time using a global registry:

```go
func init() {
    Register("SUM", sumFunc)
    Register("AVERAGE", averageFunc)
    // ... 436 more
}
```

At compile time, function names are resolved to integer IDs via `LookupFunc`. At evaluation time, `CallFunc` dispatches by ID. This design allows:

- **Fast dispatch:** function calls use integer lookup, not string comparison
- **Compact bytecode:** function IDs are small integers encoded in the instruction operand
- **Extensibility:** external packages can register new functions or override existing ones

```
Function Registry Lifecycle:

  ┌──────── init() time ────────┐   ┌──── compile time ────┐   ┌── eval time ──┐
  │                             │   │                       │   │               │
  │ Register("SUM",   sumFn)    │   │ LookupFunc("SUM")     │   │ CallFunc(0,   │
  │ Register("AVG",   avgFn)    │   │   → returns ID: 0     │   │   args, ctx)  │
  │ Register("IF",    ifFn)     │   │                       │   │   → sumFn()   │
  │ ...                         │   │ Emits: OpCall 0       │   │               │
  │                             │   │                       │   │               │
  │ Registry:                   │   └───────────────────────┘   └───────────────┘
  │ ┌────┬──────┬───────────┐   │
  │ │ ID │ Name │ Function  │   │
  │ ├────┼──────┼───────────┤   │
  │ │  0 │ SUM  │ sumFn     │   │
  │ │  1 │ AVG  │ avgFn     │   │
  │ │  2 │ IF   │ ifFn      │   │
  │ │... │ ...  │ ...       │   │
  │ │437 │ ...  │ ...       │   │
  │ └────┴──────┴───────────┘   │
  └─────────────────────────────┘
```

---

## Dependency Tracking and Incremental Recalculation

One of werkbook's more sophisticated features is its dependency graph, which enables incremental recalculation. When you change a cell's value, only the formulas that depend on that cell (directly or transitively) need to be recalculated — not every formula in the workbook.

### The Dependency Graph

The `DepGraph` (`formula/depgraph.go`) maintains bidirectional edges:

```go
type DepGraph struct {
    // forward: formula cell → cells it reads
    dependsOn  map[QualifiedCell]map[QualifiedCell]bool
    // reverse: data cell → formula cells that read it
    dependents map[QualifiedCell]map[QualifiedCell]bool
    // range subscriptions for containment checks
    rangeSubs  []rangeSub
}
```

There are two types of dependencies:

1. **Point dependencies:** Cell A5 contains `=A1+A2`. The graph records that A5 depends on A1 and A2. If either changes, A5 needs recalculation.

2. **Range subscriptions:** Cell B1 contains `=SUM(A1:A100)`. Rather than creating 100 individual point dependencies, the graph stores a range subscription. When any cell is modified, the graph checks if it falls within any subscribed range.

```
Dependency Graph Example:

  Spreadsheet:                    Dependency Graph:
  ┌─────┬─────┬──────────────┐
  │     │  A  │      B       │    Forward edges (dependsOn):
  ├─────┼─────┼──────────────┤      B3 → {A1, A2}
  │  1  │ 10  │              │      B4 → {range A1:A3}
  ├─────┼─────┼──────────────┤      B5 → {B3, B4}
  │  2  │ 20  │              │
  ├─────┼─────┼──────────────┤    Reverse edges (dependents):
  │  3  │ 30  │  =A1+A2      │      A1 → {B3}        ┐
  ├─────┼─────┼──────────────┤      A2 → {B3}        ├─ point deps
  │  4  │     │  =SUM(A1:A3) │      B3 → {B5}        │
  ├─────┼─────┼──────────────┤      B4 → {B5}        ┘
  │  5  │     │  =B3+B4      │
  └─────┴─────┴──────────────┘    Range subscriptions:
                                     B4 subscribes to A1:A3

  If A1 changes:
  ┌────────────────────────────────────────────────┐
  │  A1 modified                                   │
  │   ├── point dep → B3 marked dirty              │
  │   │                 └── point dep → B5 dirty   │
  │   └── range sub → B4 marked dirty (A1 ∈ A1:A3)│
  │                     └── point dep → B5 dirty   │
  └────────────────────────────────────────────────┘
```

### Registration

When a formula is compiled, its references are extracted and registered in the dependency graph:

```go
func (g *DepGraph) Register(formulaCell QualifiedCell, owningSheet string,
    refs []CellAddr, ranges []RangeAddr) {
    // Remove old edges first (handles formula changes)
    g.Unregister(formulaCell)

    // Record point dependencies
    for _, ref := range refs {
        // ... build forward and reverse edges
    }

    // Record range subscriptions
    for _, rng := range ranges {
        // ... store range subscription
    }
}
```

### Invalidation

When a cell's value changes via `Sheet.SetValue()`:

```
1. Update the cell's value
2. Increment File.calcGen
3. Query DepGraph for all transitive dependents
4. Mark each dependent cell as dirty (cell.dirty = true)
```

The key insight is that **formulas are not immediately recalculated**. They're just marked dirty. The actual recalculation happens lazily when `GetValue()` is called on a formula cell:

```
GetValue("A5"):
  if cell.formula != "" && (cell.dirty || cell.cachedGen != file.calcGen):
    result = evaluateFormula(cell)
    cell.value = result
    cell.cachedGen = file.calcGen
    cell.dirty = false
  return cell.value
```

This lazy approach means that if you change 1,000 cells in a loop, the recalculation cost is paid only for the formulas you actually read afterward — not for every intermediate state.

```
Lazy vs. Eager Recalculation:

Eager (what werkbook does NOT do):
  SetValue(A1) → recalc B1, C1, D1, E1, F1    ← wasted work
  SetValue(A2) → recalc B1, C1, D1, E1, F1    ← wasted work
  SetValue(A3) → recalc B1, C1, D1, E1, F1    ← wasted work
  GetValue(F1) → return cached                  Total: 15 recalculations

Lazy (what werkbook DOES):
  SetValue(A1) → mark B1,C1,D1,E1,F1 dirty    ← O(1) per dependent
  SetValue(A2) → mark B1,C1,D1,E1,F1 dirty    ← already dirty, no-op
  SetValue(A3) → mark B1,C1,D1,E1,F1 dirty    ← already dirty, no-op
  GetValue(F1) → recalc chain: B1→C1→D1→E1→F1  Total: 5 recalculations
                                                         ▲
                                                    3x fewer!
```

### Circular Reference Detection

Circular references (A1 = B1, B1 = A1) are detected at evaluation time using the `evaluating` map on the `File` struct:

```go
if f.evaluating[cellKey] {
    return ErrorValue("#REF!")
}
f.evaluating[cellKey] = true
defer delete(f.evaluating, cellKey)
// ... proceed with evaluation
```

This is simple and effective. The `evaluating` map acts as a call stack: if we encounter a cell that's already on the stack, we've found a cycle.

```
Circular Reference Detection:

  A1 = B1 + 1
  B1 = A1 + 1

  Evaluating A1:
    evaluating = {A1: true}
    │
    ├── needs B1 → evaluate B1
    │     evaluating = {A1: true, B1: true}
    │     │
    │     ├── needs A1 → evaluate A1
    │     │     A1 ∈ evaluating? YES!
    │     │     └── return #REF! ✗ cycle detected
    │     │
    │     └── B1 = #REF!
    │
    └── A1 = #REF!
```

---

## Date Handling: The 1900 Leap Year Bug

Spreadsheet date handling deserves its own section because it involves one of computing's most famous compatibility bugs.

Excel stores dates as serial numbers — the number of days since a base date. But there are two date systems:

- **1900 system (default):** Day 1 = January 1, 1900
- **1904 system (Mac legacy):** Day 1 = January 2, 1904

The 1900 system has a deliberate bug inherited from Lotus 1-2-3: it treats 1900 as a leap year (it wasn't). February 29, 1900 is serial number 60, even though that date never existed. This means:

- Serial numbers 1–59 (Jan 1 to Feb 28, 1900) are off by zero days
- Serial numbers 60+ are off by one day compared to the mathematically correct calculation

Werkbook faithfully reproduces this bug, because compatibility with Excel is more important than mathematical correctness:

```
1900 Date System — The Leap Year Bug:

Serial#   Excel says:        Reality:           Notes
──────────────────────────────────────────────────────────────
   1      Jan 1, 1900        Jan 1, 1900        ✓ correct
   2      Jan 2, 1900        Jan 2, 1900        ✓ correct
  ...         ...                ...
  59      Feb 28, 1900       Feb 28, 1900       ✓ correct
  60      Feb 29, 1900       ██ NEVER EXISTED   ✗ phantom date!
  61      Mar 1, 1900        Feb 29... wait,    ← off by one
                              actually Mar 1         from here on
  ...         ...                ...
44927     Dec 31, 2022       Dec 31, 2022       ✓ (bug cancels out)

1904 Date System (Mac):

Serial#   Date               Notes
──────────────────────────────────────────────────────────────
   0      Jan 1, 1904        No leap year bug
   1      Jan 2, 1904        Clean and correct
  ...         ...
```

The `timeToSerialForDateSystem()` function handles the conversion between Go's `time.Time` and Excel serial numbers, accounting for the appropriate date system and the leap year bug.

---

## Structured References and Tables

Werkbook supports Excel's structured reference syntax, which lets formulas refer to table columns by name:

```
=SUM(Sales[Revenue])           # Sum the Revenue column of Sales table
=Sales[#Headers]               # Reference the header row
=Sales[[#Data],[Revenue]]      # Data rows of Revenue column
=Sales[@Revenue]               # Current row's Revenue value
```

Tables are defined with column names, a reference range, and optional features like totals rows:

```go
type Table struct {
    Name      string
    SheetName string
    Ref       string
    Columns   []string
    // ...
}
```

During formula expansion (before compilation), structured references are resolved to concrete cell ranges. This expansion happens transparently — the formula engine works with regular cell references after expansion.

```
Structured Reference Expansion:

  Table "Sales" on Sheet1, range A1:D100:

  ┌─────────┬────────┬─────────┬───────┐
  │ Product │ Region │ Revenue │ Cost  │  ← row 1 (#Headers)
  ├─────────┼────────┼─────────┼───────┤
  │ Widget  │ North  │  1200   │  800  │  ← row 2
  │ Gadget  │ South  │  3400   │ 2100  │       (#Data rows)
  │ ...     │ ...    │  ...    │  ...  │
  │ Gizmo   │ East   │   900   │  600  │  ← row 100
  └─────────┴────────┴─────────┴───────┘
       A         B        C        D

  Formula                     Expands to
  ─────────────────────────────────────────────
  Sales[Revenue]          →   C2:C100
  Sales[#Headers]         →   A1:D1
  Sales[[#Data],[Cost]]   →   D2:D100
  Sales[@Revenue]         →   C{current_row}
  SUM(Sales[Revenue])     →   SUM(C2:C100)
```

---

## The CLI: `wb`

Werkbook ships with a command-line tool called `wb` that exposes the library's functionality to shell scripts, pipelines, and AI agents.

### Commands

```bash
wb info file.xlsx                        # Sheet names, dimensions, metadata
wb read file.xlsx                        # Read cell values
wb read file.xlsx --range A1:D10         # Read a specific range
wb read file.xlsx --format json          # JSON output
wb read file.xlsx --format csv           # CSV output
wb read file.xlsx --show-formulas        # Show formula text instead of values
wb edit file.xlsx --patch '[...]'        # Modify cells with a JSON patch
wb create new.xlsx --spec '{...}'        # Create a new workbook from a spec
wb calc file.xlsx                        # Force recalculate all formulas
wb dep file.xlsx                         # Show formula dependency graph
wb formula list                          # List all supported functions
```

### Agent Mode

The CLI has an `--mode agent` flag that wraps all output in a structured JSON envelope, making it easy for AI agents and automated pipelines to parse:

```json
{
  "ok": true,
  "command": "read",
  "data": { "sheets": [...] },
  "meta": {
    "schema_version": "wb.v1",
    "tool_version": "dev",
    "elapsed_ms": 45
  }
}
```

This is a thoughtful design choice — the same tool serves both humans and machines:

```
Same data, different output modes:

┌─── Human mode (default) ──────────────┐   ┌─── Agent mode (--mode agent) ──────┐
│                                        │   │                                    │
│  Sheet1                                │   │  {                                 │
│  ├─ A1: Product                        │   │    "ok": true,                     │
│  ├─ A2: Widget                         │   │    "data": {                       │
│  ├─ B1: Price                          │   │      "sheets": [{                  │
│  ├─ B2: 29.99                          │   │        "name": "Sheet1",           │
│  └─ C1: =A2&" costs $"&B2             │   │        "cells": [                  │
│                                        │   │          {"ref":"A1","v":"Product"},│
│  Readable, scannable                   │   │          ...                       │
│                                        │   │        ]                           │
└────────────────────────────────────────┘   │      }]                            │
                                             │    }                               │
                                             │  }                                 │
                                             │  Parseable, automatable            │
                                             └────────────────────────────────────┘
```

### Patch Operations

The `edit` command accepts JSON patches that describe cell modifications:

```json
[
  {"cell": "A1", "value": "Hello"},
  {"cell": "B2", "value": 42},
  {"cell": "C3", "formula": "SUM(A1:B2)"}
]
```

Edits are applied atomically — either all patches succeed or the file is left unchanged.

---

## Zero External Dependencies: A Design Philosophy

One of werkbook's most notable characteristics is its complete lack of external dependencies. The entire library is built on Go's standard library:

- `archive/zip` for ZIP archive handling
- `encoding/xml` for XML parsing and serialization
- `math` for numerical computations
- `strconv` for string/number conversions
- `time` for date handling
- `iter` for Go's range-over-func support
- `fmt`, `strings`, `sort`, `io` for utilities

This is a deliberate design choice with real benefits:

```
Dependency tree comparison:

  Typical XLSX library:              Werkbook:

  my-app                             my-app
  └── xlsx-lib                       └── werkbook
      ├── xml-parser v2.3                 └── (Go stdlib only)
      │   └── encoding-utils v1.1
      ├── zip-handler v4.0                    That's it.
      │   ├── compress-lib v3.2
      │   └── io-utils v2.0
      ├── formula-engine v1.5
      │   ├── math-ext v2.1
      │   └── parser-combinators v3.0
      └── date-utils v1.8

  12 packages to audit              0 packages to audit
  12 packages that can break        0 packages that can break
  12 packages to keep updated       0 packages to keep updated
```

1. **No supply chain risk.** No transitive dependencies means no risk of a dependency being compromised, abandoned, or introducing breaking changes.

2. **Easy vendoring.** The library can be vendored with zero additional effort.

3. **Fast compilation.** No dependency graph to resolve, no extra packages to download.

4. **Predictable behavior.** Every line of code that executes is either in werkbook or in Go's standard library, both of which you can read and reason about.

The tradeoff is that werkbook has to implement everything itself — XML parsing strategies, number formatting, all 438 formula functions. But the result is a library that's fully self-contained and under complete control.

---

## Testing: Exhaustive by Design

Werkbook's test suite is massive — over 55 test files with extensive coverage, particularly for the formula engine. The testing strategy is multi-layered:

### Formula Function Tests

Each of the 438+ functions has its own test cases, often hundreds per function. These tests use mock resolvers to evaluate formulas in isolation:

```go
// Pseudocode for typical formula test
resolver := mockResolver{
    "A1": 10,
    "A2": 20,
    "A3": 30,
}
result := evalFormula("SUM(A1:A3)", resolver)
assert(result == 60)
```

Tests cover normal operation, edge cases, error conditions, type coercion, and — critically — **Excel parity**. The goal is not just mathematical correctness but behavioral compatibility with Excel.

### Roundtrip Tests

Roundtrip tests verify that data survives a full create → save → load cycle:

```go
// Create workbook, set values and formulas
wb := werkbook.New()
// ... populate cells
wb.SaveAs("test.xlsx")

// Load it back
wb2, _ := werkbook.Open("test.xlsx")
// ... verify all values match
```

This catches subtle serialization bugs — missing XML attributes, incorrect escaping, lost styles, broken shared string table references.

### Integration Tests

Integration tests exercise the full stack: formulas that reference other formulas, cross-sheet references, multi-step dependency chains, and circular reference detection.

### Style Preservation Tests

These verify that cell styles (fonts, colors, borders, number formats) survive serialization and deserialization, and that style deduplication produces correct results.

---

## Performance Considerations

Werkbook makes several design choices that favor performance:

```
Performance Strategy Overview:

  ┌─────────────────────────────────────────────────────────────────┐
  │                    OPEN / READ                                  │
  │                                                                 │
  │  Open("file.xlsx")     GetValue("A1")    GetValue("A1") again  │
  │       │                     │                    │              │
  │       ▼                     ▼                    ▼              │
  │  Parse XML only        Compile once         Return cached      │
  │  No formula eval       Lex→Parse→Compile    value immediately  │
  │                        Cache bytecode       (calcGen match)    │
  │                        Eval via VM                             │
  │                                                                 │
  │  Cost: O(cells)        Cost: O(formula)     Cost: O(1) ✓       │
  └─────────────────────────────────────────────────────────────────┘

  ┌─────────────────────────────────────────────────────────────────┐
  │                    WRITE / MODIFY                               │
  │                                                                 │
  │  SetValue("A1", 42)   GetValue("B1")      SaveAs("out.xlsx")  │
  │       │                     │                    │              │
  │       ▼                     ▼                    ▼              │
  │  Bump calcGen          Only recalc if       Recalc dirty only  │
  │  Mark dependents       B1 is dirty or       Dedup styles       │
  │  as dirty              stale                Build SST          │
  │                                             Serialize ZIP      │
  │  Cost: O(dependents)   Cost: O(chain)       Cost: O(cells)     │
  └─────────────────────────────────────────────────────────────────┘
```

1. **Sparse data structures:** Maps instead of dense arrays mean memory usage scales with actual data, not with the dimensions of the sheet.

2. **Lazy formula evaluation:** Formulas are computed on demand, not on load. Opening a large workbook with thousands of formulas is fast because nothing is recalculated until you ask for a value.

3. **Bytecode compilation:** Formulas are compiled to bytecode once and cached. Subsequent evaluations skip the lex/parse/compile steps entirely.

4. **Constant deduplication:** The compiler deduplicates constants and cell references in the bytecode, keeping the compiled representation compact.

5. **Incremental recalculation:** The dependency graph means that changing one cell doesn't trigger recalculation of every formula — only the transitive dependents.

6. **Generation-based caching:** The `calcGen` counter provides O(1) staleness checking for cached formula values, avoiding timestamp comparisons or hash computations.

```
Generation-Based Caching:

  File.calcGen:  1        2              2              3
                 │        │              │              │
  Timeline: ─────┼────────┼──────────────┼──────────────┼─────────
                 │        │              │              │
              New()   SetValue(A1,10)  GetValue(B1)  SetValue(A1,20)
                          │              │              │
                          │         B1.cachedGen=1      │
                          │         1 ≠ 2 → stale!      │
                          │         recompute → cache    │
                          │         B1.cachedGen=2      │
                          │              │              │
                          │         GetValue(B1) again  │
                          │         B1.cachedGen=2      │
                          │         2 = 2 → fresh!      │
                          │         return cached ✓      │
                          │                              │
                          │                         B1.dirty=true
                          │                         next read will
                          │                         recompute
```

---

## Putting It All Together: A Complete Example

Let's trace through a complete workflow to see how all the pieces fit together:

```go
// 1. Create a new workbook
wb := werkbook.New()
sheet := wb.Sheet("Sheet1")

// 2. Set some values
sheet.SetValue("A1", "Product")
sheet.SetValue("A2", "Widget")
sheet.SetValue("A3", "Gadget")
sheet.SetValue("B1", "Price")
sheet.SetValue("B2", 29.99)
sheet.SetValue("B3", 49.99)
sheet.SetValue("C1", "Qty")
sheet.SetValue("C2", 100)
sheet.SetValue("C3", 50)

// 3. Set formulas
sheet.SetFormula("D1", `"Total"`)
sheet.SetFormula("D2", "B2*C2")
sheet.SetFormula("D3", "B3*C3")
sheet.SetFormula("D4", "SUM(D2:D3)")

// 4. Read a computed value
total, _ := sheet.GetValue("D4")
fmt.Println(total) // 5498.5
```

The spreadsheet in memory looks like:

```
  ┌─────────┬──────────┬─────────┬──────────────────┐
  │    A    │    B     │    C    │        D         │
  ├─────────┼──────────┼─────────┼──────────────────┤
  │Product  │ Price    │  Qty    │ ="Total"         │
  ├─────────┼──────────┼─────────┼──────────────────┤
  │Widget   │  29.99   │  100    │ =B2*C2  → ?      │
  ├─────────┼──────────┼─────────┼──────────────────┤
  │Gadget   │  49.99   │   50    │ =B3*C3  → ?      │
  ├─────────┼──────────┼─────────┼──────────────────┤
  │         │          │         │ =SUM(D2:D3) → ?  │
  └─────────┴──────────┴─────────┴──────────────────┘
```

Here's what happens under the hood when `GetValue("D4")` is called:

1. **Cell lookup:** The sheet looks up cell D4 in its sparse map.

2. **Staleness check:** D4 has a formula and its `cachedGen` doesn't match `file.calcGen`, so it needs evaluation.

3. **Compilation:** The formula `SUM(D2:D3)` is lexed, parsed into an AST, and compiled to bytecode:
   ```
   OpLoadRange  0     # Load range D2:D3
   OpCall       <SUM> # Call SUM
   ```

4. **Dependency registration:** The compiler extracts that D4 depends on the range D2:D3, and registers this in the dependency graph.

5. **Evaluation:** The VM starts executing. When it hits `OpLoadRange`, it calls the resolver, which triggers evaluation of D2 and D3 (which are also formula cells):
   - D2 (`B2*C2`): loads B2 (29.99) and C2 (100), multiplies → 2999.0
   - D3 (`B3*C3`): loads B3 (49.99) and C3 (50), multiplies → 2499.5

6. **SUM execution:** The SUM function receives the range values [2999.0, 2499.5] and returns 5498.5.

7. **Caching:** The result is cached in D4's cell, and `cachedGen` is set to the current `calcGen`.

8. **Subsequent reads:** If you call `GetValue("D4")` again without modifying any cells, the cached value is returned immediately.

Now if you modify a cell:

```go
sheet.SetValue("C2", 200) // Change Widget quantity
```

This triggers the following cascade:

```
  SetValue(C2, 200)
       │
       ▼
  calcGen: 5 → 6
       │
       ▼
  DepGraph lookup: who depends on C2?
       │
       ├──► D2 (=B2*C2)  → marked dirty
       │         │
       │         ▼
       │    who depends on D2?
       │         │
       │         └──► D4 (=SUM(D2:D3)) → marked dirty
       │
       ▼
  GetValue("D4")
       │
       ├── D4 is dirty → needs eval
       │     └── needs D2:D3
       │           ├── D2 is dirty → recompute: 29.99 * 200 = 5998.0
       │           └── D3 is clean → cached: 2499.5
       │
       └── SUM(5998.0, 2499.5) = 8497.5 ✓

  Final spreadsheet state:
  ┌─────────┬──────────┬─────────┬──────────────────────┐
  │    A    │    B     │    C    │          D           │
  ├─────────┼──────────┼─────────┼──────────────────────┤
  │Widget   │  29.99   │  200    │ =B2*C2    → 5998.0  │ ← updated
  ├─────────┼──────────┼─────────┼──────────────────────┤
  │Gadget   │  49.99   │   50    │ =B3*C3    → 2499.5  │ ← unchanged
  ├─────────┼──────────┼─────────┼──────────────────────┤
  │         │          │         │ =SUM(...) → 8497.5  │ ← updated
  └─────────┴──────────┴─────────┴──────────────────────┘
```

---

## Conclusion

Werkbook is more than a file format library. It's a complete spreadsheet engine implemented in pure Go, with a clean two-layer architecture, a bytecode-compiled formula evaluator, and an incremental recalculation system that handles dependency tracking across sheets and tables.

The design reflects a set of clear priorities: Excel compatibility over theoretical purity (the 1900 leap year bug), laziness over eagerness (formulas computed on demand), and self-containment over convenience (zero external dependencies). The result is a library that's fast, predictable, and fully self-contained — a solid foundation for any Go application that needs to work with spreadsheets as first-class data structures rather than opaque files.

Whether you're generating reports, processing financial data, building a spreadsheet-powered API, or automating workflows with the `wb` CLI, werkbook provides the tools to treat `.xlsx` files as the structured, computable documents they are.
