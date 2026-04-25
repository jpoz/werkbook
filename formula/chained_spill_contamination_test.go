package formula

import (
	"fmt"
	"testing"
)

// The tests below each reproduce one of the failures from
// testdata/chained_spill_contamination at the formula-engine level, so we can
// iterate on the architecture without rebuilding XLSX each time.
//
// Source fixtures: testdata/chained_spill_contamination/_generate.py

// ---------- test harness ----------

type contaminationResolver struct {
	cells map[string]map[CellAddr]Value
}

func (m *contaminationResolver) set(sheet string, cell string, v Value) {
	col, row, err := cellNameToColRow(cell)
	if err != nil {
		panic(err)
	}
	if m.cells == nil {
		m.cells = map[string]map[CellAddr]Value{}
	}
	if m.cells[sheet] == nil {
		m.cells[sheet] = map[CellAddr]Value{}
	}
	m.cells[sheet][CellAddr{Sheet: sheet, Col: col, Row: row}] = v
}

func (m *contaminationResolver) GetCellValue(addr CellAddr) Value {
	if sheet, ok := m.cells[addr.Sheet]; ok {
		if v, ok := sheet[addr]; ok {
			return v
		}
	}
	return EmptyVal()
}

func (m *contaminationResolver) GetRangeValues(addr RangeAddr) [][]Value {
	// Clamp to the populated extent (mimics what the real workbook resolver does).
	toRow := addr.ToRow
	toCol := addr.ToCol
	if addr.FromRow == 1 && addr.ToRow >= maxRows {
		maxRow := 0
		for ca := range m.cells[addr.Sheet] {
			if ca.Col >= addr.FromCol && ca.Col <= addr.ToCol && ca.Row > maxRow {
				maxRow = ca.Row
			}
		}
		if maxRow >= addr.FromRow {
			toRow = maxRow
		} else {
			toRow = addr.FromRow
		}
	}
	if addr.FromCol == 1 && addr.ToCol >= maxCols {
		maxCol := 0
		for ca := range m.cells[addr.Sheet] {
			if ca.Row >= addr.FromRow && ca.Row <= addr.ToRow && ca.Col > maxCol {
				maxCol = ca.Col
			}
		}
		if maxCol >= addr.FromCol {
			toCol = maxCol
		} else {
			toCol = addr.FromCol
		}
	}
	rows := make([][]Value, toRow-addr.FromRow+1)
	for r := addr.FromRow; r <= toRow; r++ {
		row := make([]Value, toCol-addr.FromCol+1)
		for c := addr.FromCol; c <= toCol; c++ {
			ca := CellAddr{Sheet: addr.Sheet, Col: c, Row: r}
			if v, ok := m.cells[addr.Sheet][ca]; ok {
				row[c-addr.FromCol] = v
			}
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

func cellNameToColRow(name string) (int, int, error) {
	col, row := 0, 0
	i := 0
	for i < len(name) && name[i] >= 'A' && name[i] <= 'Z' {
		col = col*26 + int(name[i]-'A') + 1
		i++
	}
	for i < len(name) {
		if name[i] < '0' || name[i] > '9' {
			return 0, 0, fmt.Errorf("bad cell name %q", name)
		}
		row = row*10 + int(name[i]-'0')
		i++
	}
	return col, row, nil
}

// evalFormula compiles+evaluates a formula against a resolver, returning the
// raw formula.Value so array results are preserved. `atCell` (e.g. "A1") is
// used for implicit intersection context.
func evalFormula(t *testing.T, resolver *contaminationResolver, sheet, atCell, expr string) Value {
	t.Helper()
	col, row, err := cellNameToColRow(atCell)
	if err != nil {
		t.Fatalf("bad cell %q: %v", atCell, err)
	}
	node, err := Parse(expr)
	if err != nil {
		t.Fatalf("Parse(%q): %v", expr, err)
	}
	// Prefer the alternate top-level-array bytecode when present so dynamic
	// arrays flow through rather than implicit-intersecting, but fall back to
	// the regular program for formulas whose top-level shape does not change.
	cf, err := Compile(expr, node)
	if err != nil {
		t.Fatalf("Compile(%q): %v", expr, err)
	}
	if cf.TopLevelArray != nil {
		cf = cf.TopLevelArray
	}
	ctx := &EvalContext{CurrentSheet: sheet, CurrentCol: col, CurrentRow: row}
	v, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval(%q): %v", expr, err)
	}
	return v
}

// evalScalarFormula is like evalFormula but uses the regular Compile (not
// spill-probe), so implicit intersection fires at the outer level.
func evalScalarFormula(t *testing.T, resolver *contaminationResolver, sheet, atCell, expr string) Value {
	t.Helper()
	col, row, err := cellNameToColRow(atCell)
	if err != nil {
		t.Fatalf("bad cell %q: %v", atCell, err)
	}
	node, err := Parse(expr)
	if err != nil {
		t.Fatalf("Parse(%q): %v", expr, err)
	}
	cf, err := Compile(expr, node)
	if err != nil {
		t.Fatalf("Compile(%q): %v", expr, err)
	}
	ctx := &EvalContext{CurrentSheet: sheet, CurrentCol: col, CurrentRow: row}
	v, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval(%q): %v", expr, err)
	}
	return v
}

func describeValue(v Value) string {
	switch v.Type {
	case ValueArray:
		s := "["
		for i, row := range v.Array {
			if i > 0 {
				s += "; "
			}
			s += "["
			for j, c := range row {
				if j > 0 {
					s += ", "
				}
				s += describeValue(c)
			}
			s += "]"
		}
		s += "]"
		if v.NoSpill {
			s += "{NoSpill}"
		}
		return s
	case ValueNumber:
		return fmt.Sprintf("%g", v.Num)
	case ValueString:
		return fmt.Sprintf("%q", v.Str)
	case ValueBool:
		return fmt.Sprintf("%v", v.Bool)
	case ValueError:
		return v.Err.String()
	case ValueEmpty:
		return "<empty>"
	default:
		return fmt.Sprintf("type=%d", v.Type)
	}
}

func setSrc01(r *contaminationResolver) {
	// src!A1:C1 = headers; rows are (region, product, sales)
	r.set("src", "A1", StringVal("region"))
	r.set("src", "B1", StringVal("product"))
	r.set("src", "C1", StringVal("sales"))
	rows := []struct {
		region, product string
		sales           float64
	}{
		{"north", "apple", 100},
		{"north", "banana", 200},
		{"south", "apple", 150},
		{"south", "cherry", 50},
		{"east", "banana", 300},
	}
	for i, row := range rows {
		r.set("src", fmt.Sprintf("A%d", i+2), StringVal(row.region))
		r.set("src", fmt.Sprintf("B%d", i+2), StringVal(row.product))
		r.set("src", fmt.Sprintf("C%d", i+2), NumberVal(row.sales))
	}
}

func setSrc02(r *contaminationResolver) {
	r.set("src", "A1", StringVal("name"))
	r.set("src", "B1", StringVal("score"))
	rows := []struct {
		name  string
		score float64
	}{
		{"alice", 70},
		{"bob", 90},
		{"carol", 80},
		{"dave", 60},
		{"eve", 100},
	}
	for i, row := range rows {
		r.set("src", fmt.Sprintf("A%d", i+2), StringVal(row.name))
		r.set("src", fmt.Sprintf("B%d", i+2), NumberVal(row.score))
	}
}

func setSrc03(r *contaminationResolver) {
	r.set("src", "A1", StringVal("category"))
	r.set("src", "B1", StringVal("amount"))
	rows := []struct {
		cat    string
		amount float64
	}{
		{"A", 10},
		{"B", 20},
		{"A", 15},
		{"C", 30},
		{"B", 25},
		{"A", 5},
	}
	for i, row := range rows {
		r.set("src", fmt.Sprintf("A%d", i+2), StringVal(row.cat))
		r.set("src", fmt.Sprintf("B%d", i+2), NumberVal(row.amount))
	}
}

func setData04(r *contaminationResolver) {
	r.set("data", "A1", StringVal("100,200,300,400"))
	r.set("data", "A2", StringVal("apple|banana|cherry"))
	r.set("data", "A3", StringVal("2024-01-15|2024-02-20|2024-03-25"))
	r.set("data", "A4", StringVal("a=1;b=2;c=3"))
}

func setSrc06(r *contaminationResolver) {
	r.set("src", "A1", StringVal("data"))
	for i, v := range []float64{10, 20, 30, 40, 50} {
		r.set("src", fmt.Sprintf("A%d", i+2), NumberVal(v))
	}
}

func setSrc07(r *contaminationResolver) {
	r.set("src", "A1", StringVal("key"))
	r.set("src", "B1", StringVal("val"))
	rows := []struct {
		key string
		val float64
	}{
		{"a", 10},
		{"b", 20},
		{"c", 30},
		{"d", 40},
	}
	for i, row := range rows {
		r.set("src", fmt.Sprintf("A%d", i+2), StringVal(row.key))
		r.set("src", fmt.Sprintf("B%d", i+2), NumberVal(row.val))
	}
}

// ---------- Failure 01: ROWS(FILTER(..., no-match)) ----------
// Expected (Excel cached): #VALUE!
// Current:                 #CALC!
func TestContamination01_RowsOfEmptyFilter(t *testing.T) {
	r := &contaminationResolver{}
	setSrc01(r)
	got := evalFormula(t, r, "fx", "E2",
		`ROWS(FILTER(src!A2:A6, src!A2:A6="west"))`)
	assertLookupValueEqual(t, got, ErrorVal(ErrValVALUE))
}

// ---------- Failure 02: INDEX(SORTBY(...), 0, 1) ----------
// Expected (Excel cached at anchor): "eve" (first of spilled column)
// Current:                           #VALUE! (NoSpill → #VALUE!)
func TestContamination02_IndexZeroOfSortedArray(t *testing.T) {
	r := &contaminationResolver{}
	setSrc02(r)
	got := evalFormula(t, r, "fx", "D1",
		`INDEX(SORTBY(src!A2:B6, src!B2:B6, -1), 0, 1)`)
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{StringVal("eve")},
		{StringVal("bob")},
		{StringVal("carol")},
		{StringVal("alice")},
		{StringVal("dave")},
	}})
}

// ---------- Failure 03: COUNTIF with UNIQUE criteria ----------
// Expected (Excel cached at anchor): 3 (count of first unique value)
// Current:                           0 (COUNTIF doesn't broadcast array criteria)
func TestContamination03_CountifWithArrayCriteria(t *testing.T) {
	r := &contaminationResolver{}
	setSrc03(r)
	got := evalFormula(t, r, "fx", "B1",
		`COUNTIF(src!A2:A7, UNIQUE(src!A2:A7))`)
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(2)},
		{NumberVal(1)},
	}})

	gotS := evalFormula(t, r, "fx", "C1",
		`SUMIF(src!A2:A7, UNIQUE(src!A2:A7), src!B2:B7)`)
	assertLookupValueEqual(t, gotS, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(30)},
		{NumberVal(45)},
		{NumberVal(30)},
	}})
}

// ---------- Failure 04: INDEX(TEXTSPLIT(...), 0, 2) whole-column ----------
// Expected (Excel cached at H1): 1 (first value of {"1";"2";"3"} spilled)
// Current:                       #VALUE!
func TestContamination04_IndexOfTextSplit(t *testing.T) {
	r := &contaminationResolver{}
	setData04(r)
	got := evalFormula(t, r, "fx", "H1",
		`INDEX(TEXTSPLIT(data!A4, "=", ";"), 0, 2)`)
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{StringVal("1")},
		{StringVal("2")},
		{StringVal("3")},
	}})
}

// ---------- Failure 06: INDEX with SEQUENCE row argument ----------
// Expected (Excel cached at A1): 10 (first of {10;20;30;40;50} spill)
// Current:                       #VALUE!
func TestContamination06_IndexWithSequenceArg(t *testing.T) {
	r := &contaminationResolver{}
	setSrc06(r)

	got := evalFormula(t, r, "fx", "A1",
		`INDEX(src!A2:A6, SEQUENCE(5))`)
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(10)},
		{NumberVal(20)},
		{NumberVal(30)},
		{NumberVal(40)},
		{NumberVal(50)},
	}})

	gotRev := evalFormula(t, r, "fx", "A7",
		`INDEX(src!A2:A6, 6-SEQUENCE(5))`)
	assertLookupValueEqual(t, gotRev, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(50)},
		{NumberVal(40)},
		{NumberVal(30)},
		{NumberVal(20)},
		{NumberVal(10)},
	}})

	gotBR := evalFormula(t, r, "fx", "C1",
		`BYROW(INDEX(src!A2:A6, SEQUENCE(5)), LAMBDA(r, r*2))`)
	assertLookupValueEqual(t, gotBR, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(20)},
		{NumberVal(40)},
		{NumberVal(60)},
		{NumberVal(80)},
		{NumberVal(100)},
	}})
}

// ---------- Failure 07: SINGLE / @ implicit intersection ----------
// Expected (Excel cached at F1): "yes"
// Current:                       #NAME? (SINGLE not implemented)
func TestContamination07_SingleFunction(t *testing.T) {
	r := &contaminationResolver{}
	setSrc07(r)
	got := evalFormula(t, r, "fx", "F1",
		`IF(SINGLE(FILTER(src!A2:A5, src!B2:B5>15))="b", "yes", "no")`)
	assertLookupValueEqual(t, got, StringVal("yes"))
}

func TestContamination07_AtOperator(t *testing.T) {
	r := &contaminationResolver{}
	setSrc07(r)
	got := evalFormula(t, r, "fx", "F1",
		`IF(@FILTER(src!A2:A5, src!B2:B5>15)="b", "yes", "no")`)
	assertLookupValueEqual(t, got, StringVal("yes"))
}
