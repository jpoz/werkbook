package formula

import (
	"math"
	"testing"
)

// mockResolver implements CellResolver for testing.
type mockResolver struct {
	cells map[CellAddr]Value
}

// sparseResolver mimics the workbook resolver's behavior for full-row and
// full-column references by clamping them to the populated extent.
type sparseResolver struct {
	cells map[CellAddr]Value
}

type panicRangeResolver struct {
	rangeCalls int
}

func (m *mockResolver) GetCellValue(addr CellAddr) Value {
	if v, ok := m.cells[addr]; ok {
		return v
	}
	return EmptyVal()
}

func (m *mockResolver) GetRangeValues(addr RangeAddr) [][]Value {
	rows := make([][]Value, addr.ToRow-addr.FromRow+1)
	for r := addr.FromRow; r <= addr.ToRow; r++ {
		row := make([]Value, addr.ToCol-addr.FromCol+1)
		for c := addr.FromCol; c <= addr.ToCol; c++ {
			ca := CellAddr{Sheet: addr.Sheet, Col: c, Row: r}
			if v, ok := m.cells[ca]; ok {
				row[c-addr.FromCol] = v
			}
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

func (m *sparseResolver) GetCellValue(addr CellAddr) Value {
	if v, ok := m.cells[addr]; ok {
		return v
	}
	return EmptyVal()
}

func (m *sparseResolver) GetRangeValues(addr RangeAddr) [][]Value {
	toRow := addr.ToRow
	toCol := addr.ToCol
	if addr.FromRow == 1 && addr.ToRow >= maxRows {
		toRow = max(addr.FromRow, m.maxUsedRow(addr.Sheet, addr.FromCol, addr.ToCol))
	}
	if addr.FromCol == 1 && addr.ToCol >= maxCols {
		toCol = max(addr.FromCol, m.maxUsedCol(addr.Sheet, addr.FromRow, addr.ToRow))
	}
	rows := make([][]Value, toRow-addr.FromRow+1)
	for r := addr.FromRow; r <= toRow; r++ {
		row := make([]Value, toCol-addr.FromCol+1)
		for c := addr.FromCol; c <= toCol; c++ {
			ca := CellAddr{Sheet: addr.Sheet, Col: c, Row: r}
			if v, ok := m.cells[ca]; ok {
				row[c-addr.FromCol] = v
			}
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

func (m *sparseResolver) maxUsedRow(sheet string, fromCol, toCol int) int {
	maxRow := 0
	for addr := range m.cells {
		if addr.Sheet == sheet && addr.Col >= fromCol && addr.Col <= toCol && addr.Row > maxRow {
			maxRow = addr.Row
		}
	}
	return maxRow
}

func (m *sparseResolver) maxUsedCol(sheet string, fromRow, toRow int) int {
	maxCol := 0
	for addr := range m.cells {
		if addr.Sheet == sheet && addr.Row >= fromRow && addr.Row <= toRow && addr.Col > maxCol {
			maxCol = addr.Col
		}
	}
	return maxCol
}

func (m *panicRangeResolver) GetCellValue(addr CellAddr) Value {
	return EmptyVal()
}

func (m *panicRangeResolver) GetRangeValues(addr RangeAddr) [][]Value {
	m.rangeCalls++
	return nil
}

func evalCompile(t *testing.T, formula string) *CompiledFormula {
	t.Helper()
	node, err := Parse(formula)
	if err != nil {
		t.Fatalf("Parse(%q): %v", formula, err)
	}
	cf, err := Compile(formula, node)
	if err != nil {
		t.Fatalf("Compile(%q): %v", formula, err)
	}
	return cf
}

func TestEvalArithmetic(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		{"1+2*3", 7},
		{"(1+2)*3", 9},
		{"10-3", 7},
		{"2^3", 8},
		{"2^3^2", 64}, // left-associative: (2^3)^2 = 64, not 2^(3^2) = 512
		{"10/4", 2.5},
		{"-5", -5},
		{"50%", 0.5},
		{"2+3*4-1", 13},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueNumber || got.Num != tt.want {
			t.Errorf("Eval(%q) = %v (%g), want %g", tt.formula, got.Type, got.Num, tt.want)
		}
	}
}

func TestEvalCellReference(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "A1*2")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v (%g), want 20", got.Type, got.Num)
	}
}

func TestEvalSUMRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "SUM(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("got %v (%g), want 60", got.Type, got.Num)
	}
}

func TestEvalStringConcat(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
		},
	}

	cf := evalCompile(t, `A1&" items"`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "10 items" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "10 items")
	}
}

func TestEvalIF(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IF(TRUE,"yes","no")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "yes" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "yes")
	}

	cf = evalCompile(t, `IF(FALSE,"yes","no")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "no" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "no")
	}
}

func TestEvalComparison(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "10>5")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("got %v (%v), want TRUE", got.Type, got.Bool)
	}
}

func TestEvalDivByZero(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "1/0")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("got %v (err=%v), want #DIV/0!", got.Type, got.Err)
	}
}

func TestEvalAVERAGE(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "AVERAGE(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("got %v (%g), want 15", got.Type, got.Num)
	}
}

func TestEvalMINMAX(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(15),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("MIN: got %g, want 5", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("MAX: got %g, want 15", got.Num)
	}
}

func TestEvalCOUNT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "COUNT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNT: got %g, want 2", got.Num)
	}
}

func TestEvalROUND(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "ROUND(3.14159,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || math.Abs(got.Num-3.14) > 1e-10 {
		t.Errorf("ROUND: got %g, want 3.14", got.Num)
	}
}

func TestEvalStringFunctions(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`UPPER("hello")`, "HELLO"},
		{`LOWER("HELLO")`, "hello"},
		{`TRIM("  hello   world  ")`, "hello world"},
		{`LEFT("hello",3)`, "hel"},
		{`RIGHT("hello",3)`, "llo"},
		{`MID("hello",2,3)`, "ell"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

func TestEvalLEN(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `LEN("hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("LEN: got %g, want 5", got.Num)
	}
}

func TestEvalLogical(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"AND(TRUE,TRUE)", true},
		{"AND(TRUE,FALSE)", false},
		{"OR(FALSE,TRUE)", true},
		{"OR(FALSE,FALSE)", false},
		{"NOT(TRUE)", false},
		{"NOT(FALSE)", true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestEvalIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFERROR(1/0,"err")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "err")
	}

	cf = evalCompile(t, `IFERROR(42,"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("got %v (%g), want 42", got.Type, got.Num)
	}
}

func TestEvalImplicitIntersectionFullColumn(t *testing.T) {
	// Simulate: formula in row 2 references F:F (full column).
	// In non-array formula context, F:F should be implicitly intersected
	// to a single cell at the formula's own row.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: StringVal("Header"),
			{Col: 6, Row: 2}: NumberVal(-250264),
			{Col: 6, Row: 3}: NumberVal(250264),
			{Col: 6, Row: 4}: NumberVal(-5750000),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     2,
		CurrentSheet:   "Outputs",
		IsArrayFormula: false,
	}

	// ABS(F:F) in row 2 with implicit intersection → ABS(F2) = 250264
	cf := evalCompile(t, "ABS(F:F)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 250264 {
		t.Errorf("ABS(F:F) with implicit intersection = %v (%g), want 250264", got.Type, got.Num)
	}

}

func TestEvalSUMFullRowRange(t *testing.T) {
	// SUM(5:6) should sum all values in rows 5 and 6 across all columns.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 5}: NumberVal(10),
			{Col: 2, Row: 5}: NumberVal(20),
			{Col: 3, Row: 5}: NumberVal(30),
			{Col: 1, Row: 6}: NumberVal(40),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     5,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	cf := evalCompile(t, "SUM(5:6)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 100 {
		t.Errorf("SUM(5:6) = %v (%g), want 100", got.Type, got.Num)
	}
}

func TestEvalImplicitIntersectionFullRow(t *testing.T) {
	// In a non-array formula context, 1:1 should intersect at the current column.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 2, Row: 1}: NumberVal(200),
			{Col: 3, Row: 1}: NumberVal(300),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     5,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	// ABS(1:1) in col 2 with implicit intersection → ABS(B1) = 200
	cf := evalCompile(t, "ABS(1:1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("ABS(1:1) with implicit intersection = %v (%g), want 200", got.Type, got.Num)
	}
}

func TestEvalArrayFormulaFullColumn(t *testing.T) {
	// When IsArrayFormula=true, F:F should load as a full array.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	cf := evalCompile(t, "SUM(A:A)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUM(A:A) array formula = %v (%g), want 60", got.Type, got.Num)
	}
}

func TestEvalSUMPreservesDirectFullColumnArgInScalarFormula(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: StringVal("Estimated Total Cents"),
			{Col: 5, Row: 2}: NumberVal(800000),
			{Col: 5, Row: 3}: NumberVal(0),
			{Col: 5, Row: 4}: NumberVal(0),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     1,
		CurrentSheet:   "Out - Summary",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "SUM(E:E)/100")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 8000 {
		t.Errorf("SUM(E:E)/100 = %v (%g), want 8000", got.Type, got.Num)
	}
}

func TestEvalCOUNTAPreservesDirectFullColumnArgInScalarFormula(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("ID"),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 1, Row: 3}: StringVal("b"),
			{Col: 1, Row: 4}: StringVal("c"),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     6,
		CurrentSheet:   "Out - Summary",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "MAX(COUNTA(A:A)-1,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("MAX(COUNTA(A:A)-1,0) = %v (%g), want 3", got.Type, got.Num)
	}
}

func TestEvalXLOOKUPPreservesFullColumnArgsInScalarFormula(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     4,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, `XLOOKUP("b",A:A,B:B)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf(`XLOOKUP("b",A:A,B:B) = %v (%g), want 20`, got.Type, got.Num)
	}
}

func TestEvalXMATCHPreservesFullColumnArgInScalarFormula(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     3,
		CurrentRow:     4,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "XMATCH(20,B:B)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("XMATCH(20,B:B) = %v (%g), want 2", got.Type, got.Num)
	}
}

func TestEvalFILTERPreservesFullColumnArgsInScalarFormula(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 2, Row: 1}: StringVal("drop"),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 2, Row: 2}: StringVal("keep"),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 2, Row: 3}: StringVal("keep"),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     3,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, `FILTER(A:A,B:B="keep")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("FILTER(A:A,B:B=\"keep\") type = %v, want ValueArray", got.Type)
	}
	cols := 0
	if len(got.Array) > 0 {
		cols = len(got.Array[0])
	}
	if len(got.Array) != 2 || cols != 1 || len(got.Array[1]) != 1 {
		t.Fatalf("FILTER(A:A,B:B=\"keep\") dims = %dx%d, want 2x1", len(got.Array), cols)
	}
	if got.Array[0][0].Type != ValueNumber || got.Array[0][0].Num != 200 {
		t.Fatalf("FILTER first row = %#v, want 200", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueNumber || got.Array[1][0].Num != 300 {
		t.Fatalf("FILTER second row = %#v, want 300", got.Array[1][0])
	}
}

func TestEvalPERCENTRANKPreservesFullColumnArgInScalarFormula(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "PERCENTRANK(A:A,A2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0.5 {
		t.Errorf("PERCENTRANK(A:A,A2) = %v (%g), want 0.5", got.Type, got.Num)
	}
}

func TestEvalSTANDARDIZEStillImplicitIntersectsScalarSlot(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     3,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "STANDARDIZE(A:A,AVERAGE(A:A),STDEV.S(A:A))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("STANDARDIZE(A:A,AVERAGE(A:A),STDEV.S(A:A)) = %v (%g), want 1", got.Type, got.Num)
	}
}

// ---------------------------------------------------------------------------
// coerceNum edge cases (exercised through arithmetic operations)
// ---------------------------------------------------------------------------

func TestEvalCoerceNum(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		cells   map[CellAddr]Value
		wantNum float64
		wantErr ErrorValue
		isErr   bool
	}{
		// Empty cell treated as 0
		{name: "empty+1", formula: "A1+1", wantNum: 1},
		// Bool true coerced to 1
		{name: "TRUE+0", formula: "TRUE+0", wantNum: 1},
		// Bool false coerced to 0
		{name: "FALSE+5", formula: "FALSE+5", wantNum: 5},
		// Numeric string coerced
		{name: "numeric_string", formula: `"123"+0`, wantNum: 123},
		{name: "numeric_string_float", formula: `"3.14"+0`, wantNum: 3.14},
		// Empty string produces #VALUE! (not coerced to 0)
		{name: "empty_string", formula: `""+0`, isErr: true, wantErr: ErrValVALUE},
		// Non-numeric string produces #VALUE!
		{name: "non_numeric_string", formula: `"abc"+0`, isErr: true, wantErr: ErrValVALUE},
		// Error propagates through arithmetic
		{name: "error_propagation_add", formula: `#N/A+1`, isErr: true, wantErr: ErrValNA},
		{name: "error_propagation_mul", formula: `#DIV/0!*2`, isErr: true, wantErr: ErrValDIV0},
		// Large numbers
		{name: "large_number", formula: "1e300+1e300", wantNum: 2e300},
		// Negative numbers
		{name: "negative_arithmetic", formula: "-10+-20", wantNum: -30},
		// Chained operations with coercion
		{name: "bool_chain", formula: "TRUE+TRUE+TRUE", wantNum: 3},
		// Cell containing empty string produces #VALUE!
		{name: "cell_empty_string", formula: "A1+0", cells: map[CellAddr]Value{
			{Sheet: "", Col: 1, Row: 1}: StringVal(""),
		}, isErr: true, wantErr: ErrValVALUE},
		// Cell containing numeric string coerces to number
		{name: "cell_numeric_string", formula: "A1+0", cells: map[CellAddr]Value{
			{Sheet: "", Col: 1, Row: 1}: StringVal("5"),
		}, wantNum: 5},
		// Whitespace-padded numeric string coerces to number (whitespace is trimmed)
		{name: "cell_padded_numeric_string", formula: "A1+0", cells: map[CellAddr]Value{
			{Sheet: "", Col: 1, Row: 1}: StringVal(" 5 "),
		}, wantNum: 5},
		// Whitespace-only string produces #VALUE!
		{name: "whitespace_only_string", formula: `" "+0`, isErr: true, wantErr: ErrValVALUE},
		// Truly empty cell treated as 0 (not the same as empty string)
		{name: "truly_empty_cell", formula: "A1+0", wantNum: 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cells := tt.cells
			if cells == nil {
				cells = map[CellAddr]Value{}
			}
			resolver := &mockResolver{cells: cells}
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v (err=%v), want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %v (%g), want %g", tt.formula, got.Type, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// compareValues — exercised through comparison operators
// ---------------------------------------------------------------------------

func TestEvalCompareValues(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: BoolVal(true),
			// A4 is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// Same-type number comparisons
		{name: "num_eq", formula: "10=10", want: true},
		{name: "num_ne", formula: "10<>5", want: true},
		{name: "num_lt", formula: "5<10", want: true},
		{name: "num_le", formula: "10<=10", want: true},
		{name: "num_gt", formula: "10>5", want: true},
		{name: "num_ge", formula: "10>=10", want: true},
		// Same-type string comparisons (case-insensitive)
		{name: "str_eq_case", formula: `"Hello"="hello"`, want: true},
		{name: "str_lt", formula: `"abc"<"def"`, want: true},
		{name: "str_gt", formula: `"xyz">"abc"`, want: true},
		// Same-type bool comparisons
		{name: "bool_eq", formula: "TRUE=TRUE", want: true},
		{name: "bool_ne", formula: "TRUE<>FALSE", want: true},
		{name: "bool_order", formula: "FALSE<TRUE", want: true},
		// Empty cell = 0
		{name: "empty_eq_zero", formula: "A4=0", want: true},
		// Cross-type comparisons (via typeRank: error < number < string < bool)
		{name: "num_lt_str", formula: `10<"hello"`, want: true},
		{name: "str_lt_bool", formula: `"hello"<TRUE`, want: true},
		// Negative numbers
		{name: "negative_lt", formula: "-5<0", want: true},
		{name: "negative_gt", formula: "0>-10", want: true},
		// Decimal comparisons
		{name: "decimal_eq", formula: "0.1+0.2=0.3", want: true},  // relative epsilon handles this
		{name: "third_times_3", formula: "(1/3*3)=1", want: true}, // audit: must be TRUE
		{name: "decimal_lt", formula: "0.1<0.2", want: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueBool || got.Bool != tt.want {
				t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}
}

func TestArrayElementTrimmedRangeOriginUsesBlankForMissingLogicalCells(t *testing.T) {
	trimmed := trimmedRangeValue([][]Value{
		{NumberVal(10)},
	}, 1, 1, 1, 3)

	assertLookupValueEqual(t, ArrayElement(trimmed, 0, 0), NumberVal(10))
	assertLookupValueEqual(t, ArrayElement(trimmed, 1, 0), EmptyVal())

	got := ArrayElement(trimmed, 3, 0)
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Fatalf("ArrayElement(out of bounds) = %#v, want #N/A", got)
	}
}

func TestBinaryArithTrimmedRangeOriginUsesLogicalBounds(t *testing.T) {
	trimmed := trimmedRangeValue([][]Value{
		{NumberVal(10)},
	}, 1, 1, 1, 3)

	got := binaryArith(trimmed, NumberVal(0), func(an, bn float64) Value {
		return NumberVal(an + bn)
	})

	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(10)},
		{NumberVal(0)},
		{NumberVal(0)},
	}})
	if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Fatalf("binaryArith RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
	}
}

func TestBinaryCompareTrimmedRangeOriginUsesLogicalBounds(t *testing.T) {
	trimmed := trimmedRangeValue([][]Value{
		{StringVal("keep")},
	}, 1, 1, 1, 3)

	got := binaryCompare(trimmed, StringVal("keep"), func(c int) bool { return c == 0 })

	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{BoolVal(true)},
		{BoolVal(false)},
		{BoolVal(false)},
	}})
	if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Fatalf("binaryCompare RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
	}
}

func TestLiftUnaryTrimmedRangeOriginUsesLogicalBounds(t *testing.T) {
	trimmed := trimmedRangeValue([][]Value{
		{NumberVal(-10)},
	}, 1, 1, 1, 3)

	got := LiftUnary(trimmed, func(v Value) Value {
		n, e := CoerceNum(v)
		if e != nil {
			return *e
		}
		return NumberVal(math.Abs(n))
	})

	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(10)},
		{NumberVal(0)},
		{NumberVal(0)},
	}})
	if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Fatalf("LiftUnary RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
	}
}

// ---------------------------------------------------------------------------
// isTruthy — exercised through IF
// ---------------------------------------------------------------------------

func TestEvalIsTruthy(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: StringVal(""),
			{Col: 1, Row: 4}: StringVal("x"),
			// A5 is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    string // "yes" or "no"
	}{
		{name: "bool_true", formula: `IF(TRUE,"yes","no")`, want: "yes"},
		{name: "bool_false", formula: `IF(FALSE,"yes","no")`, want: "no"},
		{name: "num_zero", formula: `IF(A1,"yes","no")`, want: "no"},
		{name: "num_nonzero", formula: `IF(A2,"yes","no")`, want: "yes"},
		{name: "str_empty", formula: `IF(A3,"yes","no")`, want: "#VALUE!"},
		{name: "str_nonempty", formula: `IF(A4,"yes","no")`, want: "#VALUE!"},
		{name: "empty_cell", formula: `IF(A5,"yes","no")`, want: "no"},
		{name: "num_negative", formula: `IF(-1,"yes","no")`, want: "yes"},
		{name: "num_fraction", formula: `IF(0.001,"yes","no")`, want: "yes"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			gotStr := ValueToString(got)
			if gotStr != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, gotStr, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// valueToString — exercised through the & (concat) operator
// ---------------------------------------------------------------------------

func TestEvalValueToString(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: BoolVal(false),
			// A4 empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "number_to_str", formula: `A1&""`, want: "42"},
		{name: "bool_true_to_str", formula: `A2&""`, want: "TRUE"},
		{name: "bool_false_to_str", formula: `A3&""`, want: "FALSE"},
		{name: "empty_to_str", formula: `A4&"x"`, want: "x"},
		{name: "string_concat", formula: `"hello"&" "&"world"`, want: "hello world"},
		{name: "float_to_str", formula: `3.14&""`, want: "3.14"},
		// error concat cases moved to TestEvalConcatErrorPropagation
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Error propagation in concat — errors must propagate, not stringify
// ---------------------------------------------------------------------------

func TestEvalConcatErrorPropagation(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{name: "DIV0_left", formula: `#DIV/0!&""`, wantErr: ErrValDIV0},
		{name: "NA_left", formula: `#N/A&""`, wantErr: ErrValNA},
		{name: "NAME_left", formula: `#NAME?&""`, wantErr: ErrValNAME},
		{name: "NULL_left", formula: `#NULL!&""`, wantErr: ErrValNULL},
		{name: "NUM_left", formula: `#NUM!&""`, wantErr: ErrValNUM},
		{name: "REF_left", formula: `#REF!&""`, wantErr: ErrValREF},
		{name: "VALUE_left", formula: `#VALUE!&""`, wantErr: ErrValVALUE},
		{name: "DIV0_right", formula: `"test"&#DIV/0!`, wantErr: ErrValDIV0},
		{name: "DIV0_from_expr", formula: `1/0&"test"`, wantErr: ErrValDIV0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Fatalf("Eval(%q) = type %v, want error", tt.formula, got.Type)
			}
			if got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = error %v, want %v", tt.formula, got.Err, tt.wantErr)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Number formatting in concat — uses 15 significant digits
// ---------------------------------------------------------------------------

func TestEvalConcatNumberFormatting(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "one_third", formula: `1/3&""`, want: "0.333333333333333"},
		{name: "small_number", formula: `0.000001&""`, want: "0.000001"},
		{name: "large_number", formula: `1000000000000000&""`, want: "1000000000000000"},
		{name: "normal_number", formula: `3.14&""`, want: "3.14"},
		{name: "integer", formula: `42&""`, want: "42"},
		{name: "zero", formula: `0&""`, want: "0"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q (type %v), want %q", tt.formula, got.Str, got.Type, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Division and power edge cases
// ---------------------------------------------------------------------------

func TestEvalDivisionEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("div_by_zero", func(t *testing.T) {
		cf := evalCompile(t, "1/0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})
	t.Run("zero_div_zero", func(t *testing.T) {
		cf := evalCompile(t, "0/0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})
	t.Run("large_div_overflow", func(t *testing.T) {
		cf := evalCompile(t, "1e300/1e-300")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || !math.IsInf(got.Num, 1) {
			t.Errorf("got %v (%g), want +Inf", got.Type, got.Num)
		}
	})
	t.Run("negative_div", func(t *testing.T) {
		cf := evalCompile(t, "-10/3")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-(-10.0/3.0)) > 1e-10 {
			t.Errorf("got %g, want %g", got.Num, -10.0/3.0)
		}
	})
	t.Run("power_zero", func(t *testing.T) {
		cf := evalCompile(t, "0^0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %g, want 1", got.Num)
		}
	})
	t.Run("power_negative_int", func(t *testing.T) {
		cf := evalCompile(t, "(-2)^3")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -8 {
			t.Errorf("got %g, want -8", got.Num)
		}
	})
	t.Run("power_fractional", func(t *testing.T) {
		cf := evalCompile(t, "4^0.5")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})
}

// ---------------------------------------------------------------------------
// Unary negation and percent with various types
// ---------------------------------------------------------------------------

func TestEvalUnaryEdgeCases(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: StringVal("5"),
		},
	}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
		wantErr ErrorValue
	}{
		{name: "negate_bool", formula: "-A1", wantNum: -1},
		{name: "negate_numeric_string", formula: "-A2", wantNum: -5},
		{name: "percent_100", formula: "100%", wantNum: 1},
		{name: "percent_50", formula: "50%", wantNum: 0.5},
		{name: "negate_zero", formula: "-0", wantNum: 0},
		{name: "double_negate", formula: "--5", wantNum: 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Error propagation through all binary operators
// ---------------------------------------------------------------------------

func TestEvalErrorPropagation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValNA),
		},
	}

	ops := []struct {
		name string
		expr string
	}{
		{name: "add_left", expr: "A1+1"},
		{name: "add_right", expr: "1+A1"},
		{name: "sub", expr: "A1-1"},
		{name: "mul", expr: "A1*2"},
		{name: "div", expr: "A1/2"},
		{name: "pow", expr: "A1^2"},
		{name: "neg", expr: "-A1"},
	}

	for _, tt := range ops {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.expr, err)
			}
			if got.Type != ValueError || got.Err != ErrValNA {
				t.Errorf("Eval(%q) = %v (err=%v), want #N/A", tt.expr, got.Type, got.Err)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// IF with two args (missing else), NOT edge cases
// ---------------------------------------------------------------------------

func TestEvalIFTwoArgs(t *testing.T) {
	resolver := &mockResolver{}

	// IF(FALSE, "yes") with no third arg should return FALSE
	cf := evalCompile(t, `IF(FALSE,"yes")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool != false {
		t.Errorf("IF(FALSE,yes) = %v, want FALSE", got)
	}
}

func TestEvalNOTEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"NOT(1)", false},
		{"NOT(0)", true},
		{`NOT("")`, true},
		{`NOT("x")`, false},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// Large number and decimal precision arithmetic
// ---------------------------------------------------------------------------

func TestEvalLargeNumberArithmetic(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		check   func(Value) bool
	}{
		{
			name:    "max_float_mul",
			formula: "1e308*1",
			check:   func(v Value) bool { return v.Type == ValueNumber && v.Num == 1e308 },
		},
		{
			name:    "overflow_to_inf",
			formula: "1e308*10",
			check:   func(v Value) bool { return v.Type == ValueNumber && math.IsInf(v.Num, 1) },
		},
		{
			name:    "very_small",
			formula: "1e-300+0",
			check:   func(v Value) bool { return v.Type == ValueNumber && v.Num == 1e-300 },
		},
		{
			name:    "subtract_near_equal",
			formula: "1.0000000001-1",
			check: func(v Value) bool {
				return v.Type == ValueNumber && math.Abs(v.Num-0.0000000001) < 1e-15
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if !tt.check(got) {
				t.Errorf("Eval(%q) = %v (%g), check failed", tt.formula, got.Type, got.Num)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// CompareValues — empty cell type adaptation
// ---------------------------------------------------------------------------

func TestCompareValues(t *testing.T) {
	tests := []struct {
		name string
		a, b Value
		want int
	}{
		// empty vs empty
		{name: "empty_empty", a: EmptyVal(), b: EmptyVal(), want: 0},

		// empty adapts to number
		{name: "empty_vs_zero", a: EmptyVal(), b: NumberVal(0), want: 0},
		{name: "zero_vs_empty", a: NumberVal(0), b: EmptyVal(), want: 0},
		{name: "empty_vs_positive", a: EmptyVal(), b: NumberVal(5), want: -1},
		{name: "positive_vs_empty", a: NumberVal(5), b: EmptyVal(), want: 1},
		{name: "empty_vs_negative", a: EmptyVal(), b: NumberVal(-3), want: 1},
		{name: "negative_vs_empty", a: NumberVal(-3), b: EmptyVal(), want: -1},

		// empty adapts to string
		{name: "empty_vs_empty_string", a: EmptyVal(), b: StringVal(""), want: 0},
		{name: "empty_string_vs_empty", a: StringVal(""), b: EmptyVal(), want: 0},
		{name: "empty_vs_nonempty_string", a: EmptyVal(), b: StringVal("hello"), want: -1},
		{name: "nonempty_string_vs_empty", a: StringVal("hello"), b: EmptyVal(), want: 1},

		// empty adapts to bool
		{name: "empty_vs_false", a: EmptyVal(), b: BoolVal(false), want: 0},
		{name: "false_vs_empty", a: BoolVal(false), b: EmptyVal(), want: 0},
		{name: "empty_vs_true", a: EmptyVal(), b: BoolVal(true), want: -1},
		{name: "true_vs_empty", a: BoolVal(true), b: EmptyVal(), want: 1},

		// same-type comparisons (sanity)
		{name: "num_eq", a: NumberVal(10), b: NumberVal(10), want: 0},
		{name: "num_lt", a: NumberVal(3), b: NumberVal(7), want: -1},
		{name: "num_gt", a: NumberVal(7), b: NumberVal(3), want: 1},
		{name: "str_eq", a: StringVal("abc"), b: StringVal("ABC"), want: 0},
		{name: "str_lt", a: StringVal("abc"), b: StringVal("def"), want: -1},
		{name: "bool_eq_true", a: BoolVal(true), b: BoolVal(true), want: 0},
		{name: "bool_eq_false", a: BoolVal(false), b: BoolVal(false), want: 0},
		{name: "bool_false_lt_true", a: BoolVal(false), b: BoolVal(true), want: -1},
		{name: "bool_true_gt_false", a: BoolVal(true), b: BoolVal(false), want: 1},

		// cross-type (different typeRank)
		{name: "num_vs_str", a: NumberVal(0), b: StringVal(""), want: -1},
		{name: "str_vs_bool", a: StringVal(""), b: BoolVal(false), want: -1},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := CompareValues(tt.a, tt.b)
			if tt.want == 0 && got != 0 {
				t.Errorf("CompareValues(%v, %v) = %d, want 0", tt.a, tt.b, got)
			} else if tt.want < 0 && got >= 0 {
				t.Errorf("CompareValues(%v, %v) = %d, want < 0", tt.a, tt.b, got)
			} else if tt.want > 0 && got <= 0 {
				t.Errorf("CompareValues(%v, %v) = %d, want > 0", tt.a, tt.b, got)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Empty cell comparison via formulas — expected behavior
// ---------------------------------------------------------------------------

func TestEvalEmptyCellComparisons(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: StringVal(""),
			{Col: 1, Row: 3}: BoolVal(false),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 1, Row: 5}: StringVal("hello"),
			{Col: 1, Row: 6}: BoolVal(true),
			// B1 is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// empty = 0 → TRUE (empty adapts to number 0)
		{name: "empty_eq_zero", formula: "B1=A1", want: true},
		{name: "zero_eq_empty", formula: "A1=B1", want: true},
		// empty = "" → TRUE (empty adapts to string "")
		{name: "empty_eq_empty_str", formula: `B1=A2`, want: true},
		{name: "empty_str_eq_empty", formula: `A2=B1`, want: true},
		// empty = FALSE → TRUE (empty adapts to bool false)
		{name: "empty_eq_false", formula: "B1=A3", want: true},
		{name: "false_eq_empty", formula: "A3=B1", want: true},
		// empty <> positive number
		{name: "empty_ne_positive", formula: "B1=A4", want: false},
		{name: "empty_lt_positive", formula: "B1<A4", want: true},
		// empty <> non-empty string
		{name: "empty_ne_nonempty_str", formula: `B1=A5`, want: false},
		{name: "empty_lt_nonempty_str", formula: `B1<A5`, want: true},
		// empty <> TRUE
		{name: "empty_ne_true", formula: "B1=A6", want: false},
		{name: "empty_lt_true", formula: "B1<A6", want: true},
		// empty = empty
		{name: "empty_eq_empty", formula: "B1=B2", want: true},
		// empty comparisons with literals
		{name: "empty_eq_zero_lit", formula: "B1=0", want: true},
		{name: "empty_eq_empty_str_lit", formula: `B1=""`, want: true},
		{name: "empty_eq_false_lit", formula: "B1=FALSE", want: true},
		{name: "empty_ne_true_lit", formula: "B1=TRUE", want: false},
		{name: "empty_ne_one_lit", formula: "B1=1", want: false},
		{name: "empty_ne_str_lit", formula: `B1="x"`, want: false},
		// inequality operators
		{name: "empty_le_zero", formula: "B1<=0", want: true},
		{name: "empty_ge_zero", formula: "B1>=0", want: true},
		{name: "empty_le_false", formula: "B1<=FALSE", want: true},
		{name: "empty_ge_false", formula: "B1>=FALSE", want: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueBool || got.Bool != tt.want {
				t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Empty cell handling in various contexts
// ---------------------------------------------------------------------------

func TestEvalEmptyCellArithmetic(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			// A2, A3 are empty
		},
	}

	tests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{name: "empty_add", formula: "A1+A2", wantNum: 5},
		{name: "empty_mul", formula: "A1*A2", wantNum: 0},
		{name: "empty_sub", formula: "A2-A1", wantNum: -5},
		{name: "sum_with_empty", formula: "SUM(A1:A3)", wantNum: 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tt.wantNum {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Array literal evaluation
// ---------------------------------------------------------------------------

func TestEvalArrayLiteral(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "SUM({1,2,3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 6 {
		t.Errorf("SUM({1,2,3}) = %g, want 6", got.Num)
	}
}

// ---------------------------------------------------------------------------
// Array binary operations — SUM(range*range)
// ---------------------------------------------------------------------------

func TestEvalArrayBinaryOps(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 3, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(3),
			{Col: 3, Row: 2}: NumberVal(4),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{name: "SUM(range*range)", formula: "SUM(B1:C1*B2:C2)", want: 11}, // 1*3 + 2*4
		{name: "SUM(range+range)", formula: "SUM(B1:C1+B2:C2)", want: 10}, // (1+3) + (2+4)
		{name: "SUM(range-range)", formula: "SUM(B1:C1-B2:C2)", want: -4}, // (1-3) + (2-4)
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, ctx)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("Eval(%q) = %v (%g), want %g", tt.formula, got.Type, got.Num, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// IFNA
// ---------------------------------------------------------------------------

func TestEvalIFNA(t *testing.T) {
	resolver := &mockResolver{}

	// #N/A should be caught
	cf := evalCompile(t, `IFNA(#N/A,"fallback")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "fallback" {
		t.Errorf("IFNA(#N/A,...) = %v, want fallback", got)
	}

	// Non-#N/A error should pass through
	cf = evalCompile(t, `IFNA(#DIV/0!,"fallback")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("IFNA(#DIV/0!,...) = %v, want #DIV/0!", got)
	}

	// Normal value passes through
	cf = evalCompile(t, `IFNA(42,"fallback")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42,...) = %v, want 42", got)
	}
}

func TestEvalErrorWrappersRespectImplicitIntersectionInScalarContext(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(11),
			{Col: 2, Row: 1}: NumberVal(22),
			{Col: 1, Row: 2}: NumberVal(202),
			{Col: 1, Row: 3}: NumberVal(303),
		},
	}

	tests := []struct {
		name    string
		formula string
		ctx     *EvalContext
		want    Value
	}{
		{
			name:    "iferror_full_row",
			formula: `IFERROR(1/0,1:1)`,
			ctx: &EvalContext{
				CurrentCol:     2,
				CurrentRow:     2,
				CurrentSheet:   "Sheet1",
				IsArrayFormula: false,
			},
			want: NumberVal(22),
		},
		{
			name:    "ifna_full_row",
			formula: `IFNA(#N/A,1:1)`,
			ctx: &EvalContext{
				CurrentCol:     2,
				CurrentRow:     2,
				CurrentSheet:   "Sheet1",
				IsArrayFormula: false,
			},
			want: NumberVal(22),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, tt.ctx)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

// ---------------------------------------------------------------------------
// 3D sheet references — parse, compile, and evaluate correctly
// ---------------------------------------------------------------------------

func TestEval3DSheetRef(t *testing.T) {
	// SUM(Sheet2:Sheet5!A11) is a 3D sheet reference (multi-sheet range).
	// With SheetListProvider support, this should evaluate by summing
	// A11 across Sheet2, Sheet3, Sheet4, Sheet5.
	resolver := &mock3DResolver{
		sheets: []string{"test", "Sheet2", "Sheet3", "Sheet4", "Sheet5"},
		cells: map[CellAddr]Value{
			{Sheet: "Sheet2", Col: 1, Row: 11}: NumberVal(1),
			{Sheet: "Sheet3", Col: 1, Row: 11}: NumberVal(2),
			{Sheet: "Sheet4", Col: 1, Row: 11}: NumberVal(3),
			{Sheet: "Sheet5", Col: 1, Row: 11}: NumberVal(4),
		},
	}

	formulas := []struct {
		formula string
		want    float64
	}{
		{"SUM(Sheet2:Sheet5!A11)", 10},
		{"SUM('Sheet2:Sheet5'!A11)", 10},
	}

	for _, tt := range formulas {
		t.Run(tt.formula, func(t *testing.T) {
			node, err := Parse(tt.formula)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.formula, err)
			}
			cf, err := Compile(tt.formula, node)
			if err != nil {
				t.Fatalf("Compile(%q) error: %v", tt.formula, err)
			}
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q) error: %v", tt.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
			}
		})
	}
}

// mock3DResolver implements CellResolver and SheetListProvider for testing 3D refs.
type mock3DResolver struct {
	sheets []string
	cells  map[CellAddr]Value
}

func (m *mock3DResolver) GetCellValue(addr CellAddr) Value {
	if v, ok := m.cells[addr]; ok {
		return v
	}
	return EmptyVal()
}

func (m *mock3DResolver) GetRangeValues(addr RangeAddr) [][]Value {
	rows := make([][]Value, addr.ToRow-addr.FromRow+1)
	for r := addr.FromRow; r <= addr.ToRow; r++ {
		row := make([]Value, addr.ToCol-addr.FromCol+1)
		for c := addr.FromCol; c <= addr.ToCol; c++ {
			row[c-addr.FromCol] = m.GetCellValue(CellAddr{Sheet: addr.Sheet, Col: c, Row: r})
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

func (m *mock3DResolver) GetSheetNames() []string {
	return m.sheets
}

// ---------------------------------------------------------------------------
// COUNTBLANK range padding — ensures blank rows beyond MaxRow are counted
// ---------------------------------------------------------------------------

func TestEvalCOUNTBLANKPadding(t *testing.T) {
	// Simulate a sheet where only rows 1 and 3 have data in column A.
	// Rows 2, 4, and 5 are blank (not present in the resolver).
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("hello"),
			{Col: 1, Row: 3}: StringVal("world"),
		},
	}

	// COUNTBLANK(A1:A5): range spans 5 rows, rows 2/4/5 are blank → 3
	cf := evalCompile(t, "COUNTBLANK(A1:A5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COUNTBLANK(A1:A5)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTBLANK(A1:A5) = %v (%g), want 3", got.Type, got.Num)
	}

	// COUNTBLANK(A1:A3): range spans 3 rows, only row 2 is blank → 1
	cf = evalCompile(t, "COUNTBLANK(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COUNTBLANK(A1:A3)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTBLANK(A1:A3) = %v (%g), want 1", got.Type, got.Num)
	}
}

func TestEvalOversizedBoundedRangeReturnsREF(t *testing.T) {
	resolver := &panicRangeResolver{}

	cf := evalCompile(t, "SUM(A1:B524289)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(SUM(A1:B524289)): %v", err)
	}
	if resolver.rangeCalls != 0 {
		t.Fatalf("GetRangeValues called %d times, want 0", resolver.rangeCalls)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("SUM(A1:B524289) = %v, want #REF!", got)
	}
}

// ---------------------------------------------------------------------------
// Implicit intersection for bounded ranges in non-array formulas
// ---------------------------------------------------------------------------

func TestEvalImplicitIntersectionBoundedRange(t *testing.T) {
	// In a non-array formula, 1+B1:B5 should implicitly intersect B1:B5
	// at the formula's row, producing a scalar result.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(0.01),
			{Col: 2, Row: 2}: NumberVal(0.02),
			{Col: 2, Row: 3}: NumberVal(-0.01),
			{Col: 2, Row: 4}: NumberVal(0.03),
			{Col: 2, Row: 5}: NumberVal(0.015),
		},
	}

	// Formula at row 3: 1+B1:B5 should intersect to B3 = -0.01, result = 0.99
	ctx := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     3,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "1+B1:B5")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || math.Abs(got.Num-0.99) > 1e-10 {
		t.Errorf("1+B1:B5 (non-array, row 3) = %v (%g), want 0.99", got.Type, got.Num)
	}

	// GEOMEAN(1+B1:B5) at row 3 should get GEOMEAN(0.99) = 0.99
	cf = evalCompile(t, "GEOMEAN(1+B1:B5)")
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval GEOMEAN: %v", err)
	}
	if got.Type != ValueNumber || math.Abs(got.Num-0.99) > 1e-10 {
		t.Errorf("GEOMEAN(1+B1:B5) (non-array, row 3) = %v (%g), want 0.99", got.Type, got.Num)
	}

	// Same formula but as array formula should compute element-wise
	ctxArray := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     3,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	cf = evalCompile(t, "GEOMEAN(1+B1:B5)")
	got, err = Eval(cf, resolver, ctxArray)
	if err != nil {
		t.Fatalf("Eval GEOMEAN (array): %v", err)
	}
	// Array mode: 1+[0.01,0.02,-0.01,0.03,0.015] = [1.01,1.02,0.99,1.03,1.015]
	// GEOMEAN of those 5 values:
	product := 1.01 * 1.02 * 0.99 * 1.03 * 1.015
	expectedGM := math.Pow(product, 1.0/5.0)
	if got.Type != ValueNumber || math.Abs(got.Num-expectedGM) > 1e-10 {
		t.Errorf("GEOMEAN(1+B1:B5) (array) = %v (%g), want %g", got.Type, got.Num, expectedGM)
	}
}

func TestEvalSUMPRODUCTArrayContext(t *testing.T) {
	// SUMPRODUCT should force array evaluation of its arguments,
	// even in non-array formula context.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     3,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	// SUMPRODUCT(A1:A3*B1:B3) = 1*4 + 2*5 + 3*6 = 32
	cf := evalCompile(t, "SUMPRODUCT(A1:A3*B1:B3)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 32 {
		t.Errorf("SUMPRODUCT(A1:A3*B1:B3) = %v (%g), want 32", got.Type, got.Num)
	}

	// SUMPRODUCT(A1:A3,B1:B3) = same result
	cf = evalCompile(t, "SUMPRODUCT(A1:A3,B1:B3)")
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 32 {
		t.Errorf("SUMPRODUCT(A1:A3,B1:B3) = %v (%g), want 32", got.Type, got.Num)
	}
}

func TestEvalSUMPRODUCTBooleanArrayDoubleNeg(t *testing.T) {
	// SUMPRODUCT(--(A1:A5="East")) should count cells equal to "East".
	// The comparison produces a boolean array, -- converts TRUE→1 / FALSE→0,
	// then SUMPRODUCT sums the values.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("East"),
			{Col: 1, Row: 2}: StringVal("West"),
			{Col: 1, Row: 3}: StringVal("East"),
			{Col: 1, Row: 4}: StringVal("East"),
			{Col: 1, Row: 5}: StringVal("North"),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, `SUMPRODUCT(--(A1:A5="East"))`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf(`SUMPRODUCT(--(A1:A5="East")) = %v (%g), want 3`, got.Type, got.Num)
	}
}

func TestEvalSUMPRODUCTWithLiftedTextAndDateFunctions(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("2026-03-12 14:58:36 UTC"),
			{Col: 2, Row: 1}: StringVal("completed"),
			{Col: 3, Row: 1}: NumberVal(10000207),
			{Col: 1, Row: 2}: StringVal("2026-02-26 17:34:12 UTC"),
			{Col: 2, Row: 2}: StringVal("completed"),
			{Col: 3, Row: 2}: NumberVal(5000000),
			{Col: 1, Row: 3}: StringVal("2025-11-06 20:04:55 UTC"),
			{Col: 2, Row: 3}: StringVal("completed"),
			{Col: 3, Row: 3}: NumberVal(5000000),
			{Col: 1, Row: 4}: EmptyVal(),
			{Col: 2, Row: 4}: StringVal("closing"),
			{Col: 3, Row: 4}: NumberVal(5000000),
			{Col: 1, Row: 5}: StringVal("2026-02-17 15:16:38 UTC"),
			{Col: 2, Row: 5}: StringVal("rejected"),
			{Col: 3, Row: 5}: NumberVal(6970032),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     4,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	// Reduced from the Excel diff workbook case that failed with #VALUE!
	cf := evalCompile(t, `SUMPRODUCT((A1:A5<>"")*(IF(A1:A5<>"",DATEVALUE(LEFT(A1:A5,10)),0)>=DATE(2025,12,12))*(B1:B5<>"rejected")*C1:C5)/100`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 150002.07 {
		t.Errorf("reduced Excel diff SUMPRODUCT = %v (%g), want 150002.07", got.Type, got.Num)
	}
}

func TestEvalSUMPRODUCTDoesNotLeakArrayContextIntoNestedIFArgs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("completed"),
			{Col: 1, Row: 2}: StringVal("completed"),
			{Col: 1, Row: 3}: StringVal("completed"),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(10),
			{Col: 3, Row: 1}: NumberVal(0),
			{Col: 3, Row: 2}: NumberVal(5),
			{Col: 3, Row: 3}: NumberVal(0),
			{Col: 4, Row: 1}: NumberVal(1),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     5,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, `SUMPRODUCT((A1:A3="completed")*IF(B1:B3-C1:C3*D1>0,B1:B3-C1:C3*D1,0))`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Without array context, IF falls back to implicit intersection. Here the
	// formula is on row 2, so Excel picks B2/C2 and returns 0 instead of an
	// element-wise array result.
	if got.Type != ValueNumber || got.Num != 0 {
		t.Fatalf("nested IF SUMPRODUCT = %v (%g), want 0", got.Type, got.Num)
	}
}

func TestEvalSUMPRODUCTNestedIFOutsideReferencedRowsReturnsValue(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 2}: StringVal("completed"),
			{Col: 1, Row: 3}: StringVal("completed"),
			{Col: 1, Row: 4}: StringVal("completed"),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(5),
			{Col: 2, Row: 4}: NumberVal(10),
			{Col: 3, Row: 2}: NumberVal(0),
			{Col: 3, Row: 3}: NumberVal(5),
			{Col: 3, Row: 4}: NumberVal(0),
			{Col: 4, Row: 2}: NumberVal(1),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, `SUMPRODUCT((A2:A4="completed")*IF(B2:B4-C2:C4*D2>0,B2:B4-C2:C4*D2,0))`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// This mirrors fixture 31: with the formula above the referenced rows,
	// implicit intersection cannot pick a row from B2:B4/C2:C4, so Excel
	// returns #VALUE! rather than evaluating IF element-wise.
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Fatalf("nested IF SUMPRODUCT outside range = %v (%v), want #VALUE!", got.Type, got.Err)
	}
}

func TestEvalElementWiseLiftedScalarFunctions(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  StringVal("alpha"),
			{Col: 1, Row: 2}:  StringVal(""),
			{Col: 1, Row: 3}:  StringVal("xy"),
			{Col: 2, Row: 1}:  NumberVal(10),
			{Col: 2, Row: 2}:  NumberVal(1),
			{Col: 2, Row: 3}:  NumberVal(100),
			{Col: 3, Row: 1}:  NumberVal(dateSerial(2026, 3, 12)),
			{Col: 3, Row: 2}:  NumberVal(dateSerial(2026, 1, 1)),
			{Col: 3, Row: 3}:  NumberVal(dateSerial(2024, 2, 29)),
			{Col: 4, Row: 1}:  NumberVal(1),
			{Col: 4, Row: 2}:  NumberVal(10),
			{Col: 4, Row: 3}:  NumberVal(100),
			{Col: 5, Row: 1}:  NumberVal(1.24),
			{Col: 5, Row: 2}:  NumberVal(2.66),
			{Col: 5, Row: 3}:  NumberVal(-3.14159),
			{Col: 6, Row: 1}:  NumberVal(1),
			{Col: 6, Row: 2}:  NumberVal(0),
			{Col: 6, Row: 3}:  NumberVal(2),
			{Col: 7, Row: 1}:  NumberVal(10),
			{Col: 7, Row: 2}:  NumberVal(1),
			{Col: 7, Row: 3}:  NumberVal(100),
			{Col: 8, Row: 1}:  StringVal("x"),
			{Col: 8, Row: 2}:  StringVal("y"),
			{Col: 8, Row: 3}:  EmptyVal(),
			{Col: 9, Row: 1}:  NumberVal(10),
			{Col: 9, Row: 2}:  NumberVal(20),
			{Col: 9, Row: 3}:  NumberVal(30),
			{Col: 10, Row: 1}: NumberVal(2),
			{Col: 10, Row: 2}: NumberVal(0),
			{Col: 10, Row: 3}: NumberVal(5),
			{Col: 11, Row: 1}: NumberVal(1),
			{Col: 11, Row: 2}: NumberVal(10),
			{Col: 11, Row: 3}: NumberVal(100),
			{Col: 12, Row: 1}: NumberVal(0),
			{Col: 12, Row: 2}: NumberVal(1),
			{Col: 12, Row: 3}: NumberVal(-1),
		},
	}

	arrayCtx := &EvalContext{
		CurrentCol:     13,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}
	scalarCtx := &EvalContext{
		CurrentCol:     13,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	tests := []struct {
		name    string
		formula string
		ctx     *EvalContext
		want    Value
	}{
		{
			name:    "array_formula_len",
			formula: "LEN(A1:A3)",
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
				{NumberVal(0)},
				{NumberVal(2)},
			}},
		},
		{
			name:    "array_formula_day",
			formula: "DAY(C1:C3)",
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(12)},
				{NumberVal(1)},
				{NumberVal(29)},
			}},
		},
		{
			name:    "array_formula_round_with_array_decimals",
			formula: "ROUND(E1:E3,F1:F3)",
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1.2)},
				{NumberVal(3)},
				{NumberVal(-3.14)},
			}},
		},
		{
			name:    "array_formula_not",
			formula: `NOT(H1:H3="x")`,
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(false)},
				{BoolVal(true)},
				{BoolVal(true)},
			}},
		},
		{
			name:    "array_formula_iferror",
			formula: "IFERROR(I1:I3/J1:J3,9)",
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
				{NumberVal(9)},
				{NumberVal(6)},
			}},
		},
		{
			name:    "array_formula_norm_dist",
			formula: "ROUND(NORM.DIST(L1:L3,0,1,TRUE),3)",
			ctx:     arrayCtx,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(0.5)},
				{NumberVal(0.841)},
				{NumberVal(0.159)},
			}},
		},
		{
			name:    "sumproduct_len",
			formula: "SUMPRODUCT(LEN(A1:A3),B1:B3)",
			ctx:     scalarCtx,
			want:    NumberVal(250),
		},
		{
			name:    "sumproduct_day_filter",
			formula: "SUMPRODUCT((DAY(C1:C3)>=10)*D1:D3)",
			ctx:     scalarCtx,
			want:    NumberVal(101),
		},
		{
			name:    "sumproduct_round",
			formula: "SUMPRODUCT(ROUND(E1:E3,F1:F3),G1:G3)",
			ctx:     scalarCtx,
			want:    NumberVal(-299),
		},
		{
			// IFERROR evaluates its first arg in scalar context even when
			// nested in SUMPRODUCT: at CurrentRow=1 the I1:I3/J1:J3
			// expression implicit-intersects to I1/J1 = 10/2 = 5, so
			// IFERROR(5,0)=5 and SUMPRODUCT(5, K1:K3) can't line up the
			// scalar against the 3-cell K1:K3 — #VALUE!. Verified against
			// Excel in testdata/error_propagation/12_sumproduct_iferror_div.xlsx.
			name:    "sumproduct_iferror",
			formula: "SUMPRODUCT(IFERROR(I1:I3/J1:J3,0),K1:K3)",
			ctx:     scalarCtx,
			want:    ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, tt.ctx)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestEvalImplicitIntersectionRowVector(t *testing.T) {
	// Row vector implicit intersection: single-row range intersects at formula's column.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 5}: NumberVal(10),
			{Col: 2, Row: 5}: NumberVal(20),
			{Col: 3, Row: 5}: NumberVal(30),
		},
	}

	// Formula at col 2: 1+A5:C5 should intersect to B5 = 20, result = 21
	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	cf := evalCompile(t, "1+A5:C5")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 21 {
		t.Errorf("1+A5:C5 (non-array, col 2) = %v (%g), want 21", got.Type, got.Num)
	}
}

// TestEvalIFERRORImplicitIntersection verifies that IFERROR's raw range-ref
// arguments undergo legacy implicit intersection in a non-array formula
// context, matching Excel's cached results.
func TestEvalIFERRORImplicitIntersection(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: ErrorVal(ErrValNA),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: ErrorVal(ErrValDIV0),
			{Col: 1, Row: 5}: NumberVal(50),
		},
	}

	// At row 1 the implicit intersection of A1:A5 picks A1 = 10.
	ctxRow1 := &EvalContext{CurrentCol: 5, CurrentRow: 1, CurrentSheet: "Sheet1"}
	cf := evalCompile(t, `SUM(IFERROR(A1:A5, 0))`)
	got, err := Eval(cf, resolver, ctxRow1)
	if err != nil {
		t.Fatalf("Eval row 1: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("SUM(IFERROR(A1:A5, 0)) at row 1 = %v (%g), want 10", got.Type, got.Num)
	}

	// At row 2 the intersection picks A2 = #N/A, IFERROR substitutes 0, SUM = 0.
	ctxRow2 := &EvalContext{CurrentCol: 5, CurrentRow: 2, CurrentSheet: "Sheet1"}
	got, err = Eval(cf, resolver, ctxRow2)
	if err != nil {
		t.Fatalf("Eval row 2: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("SUM(IFERROR(A1:A5, 0)) at row 2 = %v (%g), want 0", got.Type, got.Num)
	}

	// With an alternative replacement the same pattern resolves to -1.
	ctxRow2b := &EvalContext{CurrentCol: 5, CurrentRow: 2, CurrentSheet: "Sheet1"}
	cf2 := evalCompile(t, `SUM(IFERROR(A1:A5, -1))`)
	got2, err := Eval(cf2, resolver, ctxRow2b)
	if err != nil {
		t.Fatalf("Eval row 2 alt: %v", err)
	}
	if got2.Type != ValueNumber || got2.Num != -1 {
		t.Errorf("SUM(IFERROR(A1:A5, -1)) at row 2 = %v (%g), want -1", got2.Type, got2.Num)
	}
}

// TestEvalRangeIntersectOperator exercises the space-intersection operator:
// a space between two references returns the rectangular overlap.
func TestEvalRangeIntersectOperator(t *testing.T) {
	cells := map[CellAddr]Value{}
	// Populate A1:D4 with row*10+col so overlap values are easy to verify.
	for r := 1; r <= 4; r++ {
		for c := 1; c <= 4; c++ {
			cells[CellAddr{Col: c, Row: r}] = NumberVal(float64(r*10 + c))
		}
	}
	resolver := &mockResolver{cells: cells}

	// SUM(A1:C3 B2:D4): overlap is B2:C3 = 22, 23, 32, 33 → 110.
	ctx := &EvalContext{CurrentCol: 7, CurrentRow: 2, CurrentSheet: "Sheet1"}
	cf := evalCompile(t, `SUM(A1:C3 B2:D4)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval SUM: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 110 {
		t.Errorf("SUM(A1:C3 B2:D4) = %v (%g), want 110", got.Type, got.Num)
	}

	// Scalar context at row 1: B2:C3 has no row-1 row → implicit intersection fails → #VALUE!.
	ctxRow1 := &EvalContext{CurrentCol: 7, CurrentRow: 1, CurrentSheet: "Sheet1"}
	cfPlain := evalCompile(t, `A1:C3 B2:D4`)
	got, err = Eval(cfPlain, resolver, ctxRow1)
	if err != nil {
		t.Fatalf("Eval plain at row 1: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("A1:C3 B2:D4 at row 1 = %v (%v), want #VALUE!", got.Type, got.Err)
	}

	// Empty intersection: A1:B2 and D4:E5 don't overlap → #NULL!.
	cfNull := evalCompile(t, `A1:B2 D4:E5`)
	got, err = Eval(cfNull, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval empty isect: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNULL {
		t.Errorf("A1:B2 D4:E5 = %v (%v), want #NULL!", got.Type, got.Err)
	}

	// SUM of empty intersection propagates #NULL!.
	cfSumNull := evalCompile(t, `SUM(A1:B2 D4:E5)`)
	got, err = Eval(cfSumNull, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval SUM of empty isect: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNULL {
		t.Errorf("SUM(A1:B2 D4:E5) = %v (%v), want #NULL!", got.Type, got.Err)
	}

	// IFERROR wrapping an empty intersection substitutes the fallback.
	cfIferr := evalCompile(t, `IFERROR(A1:B2 D4:E5, -99)`)
	got, err = Eval(cfIferr, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval IFERROR of empty isect: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -99 {
		t.Errorf("IFERROR(A1:B2 D4:E5, -99) = %v (%g), want -99", got.Type, got.Num)
	}
}

// TestEvalUnionReferenceInFunctions verifies that a parenthesized union
// reference list passed to generic functions flattens into a single array
// input (matches Excel behaviour for SUM and COUNT).
func TestEvalUnionReferenceInFunctions(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(11),
			{Col: 1, Row: 2}: NumberVal(21),
			{Col: 3, Row: 1}: NumberVal(13),
			{Col: 3, Row: 2}: NumberVal(23),
			{Col: 5, Row: 5}: NumberVal(99),
		},
	}

	ctx := &EvalContext{CurrentCol: 7, CurrentRow: 3, CurrentSheet: "Sheet1"}

	cfSum := evalCompile(t, `SUM((A1:A2, C1:C2))`)
	got, err := Eval(cfSum, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval SUM union: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 68 {
		t.Errorf("SUM((A1:A2, C1:C2)) = %v (%g), want 68", got.Type, got.Num)
	}

	cfCount := evalCompile(t, `COUNT((A1:A2, C1:C2, E5))`)
	got, err = Eval(cfCount, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval COUNT union: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("COUNT((A1:A2, C1:C2, E5)) = %v (%g), want 5", got.Type, got.Num)
	}
}
