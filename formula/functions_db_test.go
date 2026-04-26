package formula

import (
	"math"
	"testing"
)

// makeDBResolver builds a mockResolver with a database in A1:E11 style layout
// and optional criteria cells. The database argument is a [][]Value where the
// first row contains headers. Criteria are placed starting at the given cell
// address offset.
func makeDBResolver(db [][]Value, dbStartCol, dbStartRow int, crit [][]Value, critStartCol, critStartRow int) *mockResolver {
	cells := make(map[CellAddr]Value)
	for r, row := range db {
		for c, v := range row {
			cells[CellAddr{Col: dbStartCol + c, Row: dbStartRow + r}] = v
		}
	}
	for r, row := range crit {
		for c, v := range row {
			cells[CellAddr{Col: critStartCol + c, Row: critStartRow + r}] = v
		}
	}
	return &mockResolver{cells: cells}
}

// dbRange returns a range string like "A1:D5" for a database grid.
func dbRange(startCol, startRow, numCols, numRows int) string {
	return cellName(startCol, startRow) + ":" + cellName(startCol+numCols-1, startRow+numRows-1)
}

func cellName(col, row int) string {
	c := string(rune('A' + col - 1))
	return c + dbItoa(row)
}

func dbItoa(n int) string {
	s := ""
	for n > 0 {
		s = string(rune('0'+n%10)) + s
		n /= 10
	}
	if s == "" {
		return "0"
	}
	return s
}

func TestDSUM(t *testing.T) {
	// Standard test database (similar to Excel documentation):
	// | Tree    | Height | Age | Yield | Profit |
	// | Apple   | >10    | ... | ...   | ...    |
	// | Pear    | ...    | ... | ...   | ...    |
	// etc.
	//
	// Database in A5:E11 (rows 5-11, cols A-E)
	// Criteria in various locations

	// Database: columns = Tree, Height, Age, Yield, Profit
	db := [][]Value{
		{StringVal("Tree"), StringVal("Height"), StringVal("Age"), StringVal("Yield"), StringVal("Profit")},
		{StringVal("Apple"), NumberVal(18), NumberVal(20), NumberVal(14), NumberVal(105)},
		{StringVal("Pear"), NumberVal(12), NumberVal(12), NumberVal(10), NumberVal(96)},
		{StringVal("Cherry"), NumberVal(13), NumberVal(14), NumberVal(9), NumberVal(105)},
		{StringVal("Apple"), NumberVal(14), NumberVal(15), NumberVal(10), NumberVal(75)},
		{StringVal("Pear"), NumberVal(9), NumberVal(8), NumberVal(8), NumberVal(76.8)},
		{StringVal("Apple"), NumberVal(8), NumberVal(9), NumberVal(6), NumberVal(45)},
	}

	tests := []struct {
		name     string
		db       [][]Value
		crit     [][]Value
		field    string // field as string in the formula
		wantType ValueType
		wantNum  float64
		wantErr  ErrorValue
	}{
		{
			name: "basic DSUM with string field - Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225, // 105 + 75 + 45
		},
		{
			name: "DSUM with numeric field index",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // 5th column = Profit
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "DSUM with Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  172.8, // 96 + 76.8
		},
		{
			name: "multiple criteria AND - Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // 105 (h=18) + 75 (h=14)
		},
		{
			name: "multiple criteria rows OR - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  397.8, // 225 + 172.8
		},
		{
			name: "numeric comparison >10 on Height, sum Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  43, // 14+10+9+10 (heights 18,12,13,14)
		},
		{
			name: "numeric comparison <10 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  121.8, // 76.8 (h=9) + 45 (h=8)
		},
		{
			name: "numeric comparison >=14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // 105 (h=18) + 75 (h=14)
		},
		{
			name: "numeric comparison <=12 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  217.8, // 96 (h=12) + 76.8 (h=9) + 45 (h=8)
		},
		{
			name: "numeric comparison <>13 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  397.8, // all except Cherry (105) = 105+96+75+76.8+45
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14
		},
		{
			name: "text case-insensitive match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "blank criteria matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")}, // blank = match all
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8, // sum of all profits
		},
		{
			name: "no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				// no condition rows
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8,
		},
		{
			name: "no matching records returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Orange")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name:     "field name not found returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index 0 returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "wildcard * in criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")}, // matches Apple
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "wildcard ? in criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")}, // matches Pear
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  172.8,
		},
		{
			name: "mixed types in field column - only sum numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")}, // match all
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  30, // 10 + 20 (text, bool, empty ignored)
		},
		{
			name: "multiple AND criteria with text and numeric",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // Apple with age>10: age 20 (profit 105) + age 15 (profit 75)
		},
		{
			name: "complex OR criteria - Apple with height>10 OR Pear with height<10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  256.8, // Apple h>10: 105+75=180, Pear h<10: 76.8 → 256.8
		},
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				// no data rows
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "exact match with = prefix",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Apple"), NumberVal(10)},
				{StringVal("Apple Pie"), NumberVal(20)},
				{StringVal("apple"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=Apple")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  40, // "Apple" and "apple" match (case-insensitive), not "Apple Pie"
		},
		{
			name: "single record match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "negative numbers in sum field",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(5)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -25, // -10 + -20 + 5
		},
		{
			name: "sum of zeros",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(0)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "large database - 50 rows",
			db: func() [][]Value {
				rows := [][]Value{
					{StringVal("Category"), StringVal("Amount")},
				}
				for i := 1; i <= 50; i++ {
					cat := "A"
					if i%2 == 0 {
						cat = "B"
					}
					rows = append(rows, []Value{StringVal(cat), NumberVal(float64(i))})
				}
				return rows
			}(),
			crit: [][]Value{
				{StringVal("Category")},
				{StringVal("A")}, // odd numbers: 1+3+5+...+49 = 625
			},
			field:    `"Amount"`,
			wantType: ValueNumber,
			wantNum:  625,
		},
		{
			name: "boolean criteria - match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Value")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  40, // 10 + 30
		},
		{
			name: "boolean criteria - match FALSE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Value")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(false)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  20,
		},
		{
			name: "criteria header not in database - no match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(20)},
			},
			crit: [][]Value{
				{StringVal("NonExistentColumn")},
				{StringVal("A")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0, // criteria column not found -> never matches
		},
		{
			name: "criteria on different column than summed field",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  24, // Age>14: age 20 yield 14, age 15 yield 10 -> 24
		},
		{
			name: "numeric criteria value (not string comparison)",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score"), StringVal("Value")},
				{StringVal("A"), NumberVal(100), NumberVal(10)},
				{StringVal("B"), NumberVal(200), NumberVal(20)},
				{StringVal("C"), NumberVal(100), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Score")},
				{NumberVal(100)}, // exact numeric match
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  40, // 10 + 30
		},
		{
			name: "field specified as 1 (first column)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "1", // first column = Tree (text, so sum ignores it)
			wantType: ValueNumber,
			wantNum:  0, // Tree column is text, DSUM ignores text
		},
		{
			name: "field index negative returns #VALUE!",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "combined AND/OR - (Apple AND Height>10) OR (Cherry)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},  // row 1: Apple AND Height>10
				{StringVal("Cherry"), StringVal("")},     // row 2: Cherry (blank Height = match all)
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  285, // Apple h>10: 105+75=180, Cherry: 105 -> 285
		},
		{
			name: "wildcard * matches all trees starting with letter",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // contains 'e': Apple, Pear, Cherry
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8, // all trees contain 'e'
		},
		{
			name: "wildcard ? single char - Pea? matches Pear only",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Pear"), NumberVal(10)},
				{StringVal("Peas"), NumberVal(20)},
				{StringVal("Peak"), NumberVal(30)},
				{StringVal("Pearl"), NumberVal(40)}, // 5 chars, no match for Pea?
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Pea?")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  60, // Pear(10) + Peas(20) + Peak(30)
		},
		{
			name: "sum Yield where Profit > 100",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  23, // Profit>100: Apple(105,yield=14), Cherry(105,yield=9) -> 23
		},
		{
			name: "text in summed column ignored",
			db: [][]Value{
				{StringVal("Name"), StringVal("Amount")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), StringVal("N/A")},
				{StringVal("C"), NumberVal(200)},
				{StringVal("D"), StringVal("pending")},
				{StringVal("E"), NumberVal(50)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Amount"`,
			wantType: ValueNumber,
			wantNum:  350, // 100 + 200 + 50
		},
		{
			name: "three column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // Apple, Height>10, Age<20: only Apple h=14 age=15
		},
		{
			name: "criteria with <> on text - not Apple",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("<>Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  277.8, // Pear(96+76.8) + Cherry(105) = 277.8
		},
		{
			name: "decimal values in sum",
			db: [][]Value{
				{StringVal("Item"), StringVal("Price")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.75)},
				{StringVal("C"), NumberVal(3.25)},
			},
			crit: [][]Value{
				{StringVal("Item")},
				{StringVal("")},
			},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  7.5,
		},
		{
			name: "multiple OR rows with same column value",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=18")},
				{StringVal("=8")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  150, // h=18 profit 105, h=8 profit 45
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// Place database at A1
			dbRows := len(tt.db)
			dbCols := len(tt.db[0])
			// Place criteria at col G (7), row 1
			critRows := len(tt.crit)
			critCols := len(tt.crit[0])

			resolver := makeDBResolver(tt.db, 1, 1, tt.crit, 7, 1)

			formula := "DSUM(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tt.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}

			if got.Type != tt.wantType {
				t.Fatalf("DSUM type = %v, want %v (value: %+v)", got.Type, tt.wantType, got)
			}

			switch tt.wantType {
			case ValueNumber:
				if diff := got.Num - tt.wantNum; diff > 1e-9 || diff < -1e-9 {
					t.Errorf("DSUM = %g, want %g", got.Num, tt.wantNum)
				}
			case ValueError:
				if got.Err != tt.wantErr {
					t.Errorf("DSUM error = %v, want %v", got.Err, tt.wantErr)
				}
			}
		})
	}
}

func TestDSUM_WrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few arguments
	cf := evalCompile(t, "DSUM(A1:B2,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// With only 2 args the function should return an error.
	// Note: the parser may not pass exactly 2 args depending on how
	// the formula is compiled. We test the function directly too.
	_ = got

	// Direct function call with wrong arg counts
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDSum(tt.args)
			if err != nil {
				t.Fatalf("fnDSum error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDSum(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDSUM_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the summed field, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := "DSUM(A1:B4,\"Value\",G1:G2)"
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DSUM with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDSUM_FieldCaseInsensitive(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("VALUE")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	// Field name uses different case
	formula := `DSUM(A1:B3,"value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("DSUM case-insensitive field = %+v, want 30", got)
	}
}

// ---------------------------------------------------------------------------
// Helper: run a D-function table-driven test
// ---------------------------------------------------------------------------

// dbTestCase is shared by all D-function table tests.
type dbTestCase struct {
	name     string
	db       [][]Value
	crit     [][]Value
	field    string
	wantType ValueType
	wantNum  float64
	wantStr  string
	wantBool bool
	wantErr  ErrorValue
}

// runDBTests runs table-driven tests for a D-function.
func runDBTests(t *testing.T, funcName string, tests []dbTestCase) {
	t.Helper()
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dbRows := len(tt.db)
			dbCols := len(tt.db[0])
			critRows := len(tt.crit)
			critCols := len(tt.crit[0])

			resolver := makeDBResolver(tt.db, 1, 1, tt.crit, 7, 1)

			formula := funcName + "(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tt.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}

			if got.Type != tt.wantType {
				t.Fatalf("%s type = %v, want %v (value: %+v)", funcName, got.Type, tt.wantType, got)
			}

			switch tt.wantType {
			case ValueNumber:
				if diff := got.Num - tt.wantNum; diff > 1e-9 || diff < -1e-9 {
					t.Errorf("%s = %g, want %g", funcName, got.Num, tt.wantNum)
				}
			case ValueError:
				if got.Err != tt.wantErr {
					t.Errorf("%s error = %v, want %v", funcName, got.Err, tt.wantErr)
				}
			case ValueString:
				if got.Str != tt.wantStr {
					t.Errorf("%s = %q, want %q", funcName, got.Str, tt.wantStr)
				}
			case ValueBool:
				if got.Bool != tt.wantBool {
					t.Errorf("%s = %v, want %v", funcName, got.Bool, tt.wantBool)
				}
			}
		})
	}
}

func TestDSUM_TrimmedCriteriaLogicalBlankRowMatchesAll(t *testing.T) {
	db := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("Tree"), StringVal("Profit")},
			{StringVal("Apple"), NumberVal(10)},
			{StringVal("Pear"), NumberVal(20)},
		},
	}
	criteria := trimmedRangeValue([][]Value{
		{StringVal("Tree")},
		{StringVal("Apple")},
	}, 7, 1, 7, 3)

	got, err := fnDSum([]Value{db, StringVal("Profit"), criteria})
	if err != nil {
		t.Fatalf("fnDSum(trimmed criteria with logical blank row): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Fatalf("fnDSum(trimmed criteria with logical blank row) = %+v, want 30", got)
	}
}

func TestDSUM_FullColumnCriteriaIgnoresLogicalBlankTail(t *testing.T) {
	db := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("Tree"), StringVal("Profit")},
			{StringVal("Apple"), NumberVal(10)},
			{StringVal("Pear"), NumberVal(20)},
		},
	}
	criteria := trimmedRangeValue([][]Value{
		{StringVal("Tree")},
		{StringVal("Apple")},
	}, 7, 1, 7, maxRows)

	got, err := fnDSum([]Value{db, StringVal("Profit"), criteria})
	if err != nil {
		t.Fatalf("fnDSum(full-column trimmed criteria): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Fatalf("fnDSum(full-column trimmed criteria) = %+v, want 10", got)
	}
}

// standardDB is the shared test database used across D-function tests.
func standardDB() [][]Value {
	return [][]Value{
		{StringVal("Tree"), StringVal("Height"), StringVal("Age"), StringVal("Yield"), StringVal("Profit")},
		{StringVal("Apple"), NumberVal(18), NumberVal(20), NumberVal(14), NumberVal(105)},
		{StringVal("Pear"), NumberVal(12), NumberVal(12), NumberVal(10), NumberVal(96)},
		{StringVal("Cherry"), NumberVal(13), NumberVal(14), NumberVal(9), NumberVal(105)},
		{StringVal("Apple"), NumberVal(14), NumberVal(15), NumberVal(10), NumberVal(75)},
		{StringVal("Pear"), NumberVal(9), NumberVal(8), NumberVal(8), NumberVal(76.8)},
		{StringVal("Apple"), NumberVal(8), NumberVal(9), NumberVal(6), NumberVal(45)},
	}
}

// ---------------------------------------------------------------------------
// DAVERAGE
// ---------------------------------------------------------------------------

func TestDAVERAGE(t *testing.T) {
	db := standardDB()

	mixedDB := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), StringVal("text")},
		{StringVal("C"), NumberVal(20)},
		{StringVal("D"), BoolVal(true)},
		{StringVal("E"), EmptyVal()},
	}

	tests := []dbTestCase{
		{
			name: "average Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // (105+75+45)/3
		},
		{
			name: "average Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  86.4, // (96+76.8)/2
		},
		{
			name: "average all yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  57.0 / 6.0, // (14+10+9+10+8+6)/6 = 9.5
		},
		{
			name: "average with numeric criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  95.25, // (105+96+105+75)/4
		},
		{
			name: "average single matching record",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name:     "average no matching records returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average with mixed types - only numeric",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  15, // (10+20)/2
		},
		{
			name: "average with numeric field index",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  75,
		},
		{
			name: "average with AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  90, // (105+75)/2
		},
		{
			name: "average with OR criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  82.5, // (105+75+45+105)/4
		},
		{
			name:     "average empty database returns DIV/0",
			db:       [][]Value{{StringVal("Name"), StringVal("Value")}},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average all text values returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("x")},
				{StringVal("B"), StringVal("y")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average height of all trees",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  74.0 / 6.0, // (18+12+13+14+9+8)/6
		},
		{
			name:     "average field not found",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"Missing"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "average wildcard criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("P*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  86.4, // (96+76.8)/2
		},
		// --- additional comprehensive tests ---
		{
			name: "no criteria rows matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				// no condition rows → match all
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8 / 6.0, // (105+96+105+75+76.8+45)/6
		},
		{
			name: "field index 1 averages first column (text ignored → DIV/0)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "1", // first column = Tree (text), DAVERAGE skips text
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "field index out of range returns VALUE error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "field index 0 returns VALUE error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "field index negative returns VALUE error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "numeric comparison < on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  (76.8 + 45) / 2.0, // h=9 (76.8), h=8 (45)
		},
		{
			name: "numeric comparison >= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  (105 + 75) / 2.0, // h=18 (105), h=14 (75)
		},
		{
			name: "numeric comparison <= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  (96 + 76.8 + 45) / 3.0, // h=12 (96), h=9 (76.8), h=8 (45)
		},
		{
			name: "numeric comparison <> on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  397.8 / 5.0, // all except Cherry (105): 105+96+75+76.8+45=397.8
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14 → single value average = 75
		},
		{
			name: "combined AND/OR - (Apple AND Height>10) OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Cherry"), StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  285.0 / 3.0, // Apple h>10: 105+75, Cherry: 105 → (285)/3=95
		},
		{
			name: "wildcard ? single char match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")}, // matches Pear
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  86.4, // (96+76.8)/2
		},
		{
			name: "wildcard * contains pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // all trees contain 'e'
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8 / 6.0,
		},
		{
			name: "case-insensitive text criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // same as Apple: (105+75+45)/3
		},
		{
			name: "criteria not-equal text - not Apple",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("<>Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  277.8 / 3.0, // Pear(96+76.8) + Cherry(105) = 277.8 / 3
		},
		{
			name: "cross-column criteria - filter by Age, average Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  24.0 / 2.0, // Age>14: age20 yield14, age15 yield10 → 12
		},
		{
			name: "cross-column criteria - filter by Profit, average Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  (18 + 13) / 2.0, // Profit>100: Apple(h=18,p=105), Cherry(h=13,p=105) → 15.5
		},
		{
			name: "three column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14 age=15
		},
		{
			name: "negative numbers in averaged field",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(5)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -25.0 / 3.0, // (-10 + -20 + 5)/3
		},
		{
			name: "decimal values in averaged field",
			db: [][]Value{
				{StringVal("Item"), StringVal("Price")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.75)},
				{StringVal("C"), NumberVal(3.25)},
			},
			crit: [][]Value{
				{StringVal("Item")},
				{StringVal("")},
			},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  7.5 / 3.0, // (1.5+2.75+3.25)/3 = 2.5
		},
		{
			name: "average of identical values equals that value",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
				{StringVal("C"), NumberVal(42)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		{
			name: "average of zeros",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(0)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "boolean criteria match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Value")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  20, // (10+30)/2
		},
		{
			name: "criteria header not in database causes no match → DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(20)},
			},
			crit: [][]Value{
				{StringVal("NonExistentColumn")},
				{StringVal("A")},
			},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0, // no matches → zero numeric values → DIV/0
		},
		{
			name: "numeric criteria value exact match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score"), StringVal("Value")},
				{StringVal("A"), NumberVal(100), NumberVal(10)},
				{StringVal("B"), NumberVal(200), NumberVal(20)},
				{StringVal("C"), NumberVal(100), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Score")},
				{NumberVal(100)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  20, // (10+30)/2
		},
		{
			name: "exact match with = prefix excludes partial matches",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Apple"), NumberVal(10)},
				{StringVal("Apple Pie"), NumberVal(20)},
				{StringVal("apple"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=Apple")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  20, // "Apple" and "apple" (case-insensitive) → (10+30)/2=20
		},
		{
			name: "multiple OR rows with specific heights",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=18")},
				{StringVal("=8")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  (105 + 45) / 2.0, // h=18 profit 105, h=8 profit 45 → 75
		},
		{
			name: "field name not found returns VALUE error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Missing"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "complex OR three rows - Apple OR Pear OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8 / 6.0, // all 6 records match
		},
		{
			name: "average with field header case mismatch",
			db: [][]Value{
				{StringVal("Name"), StringVal("VALUE")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(20)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"value"`, // lowercase vs uppercase header
			wantType: ValueNumber,
			wantNum:  15, // (10+20)/2
		},
	}

	runDBTests(t, "DAVERAGE", tests)
}

func TestDAVERAGE_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the averaged field, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DAVERAGE(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DAVERAGE with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDAVERAGE_WrongArgCount(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDAverage(tt.args)
			if err != nil {
				t.Fatalf("fnDAverage error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDAverage(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// DCOUNT
// ---------------------------------------------------------------------------

func TestDCOUNT(t *testing.T) {
	db := standardDB()

	mixedDB := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), StringVal("text")},
		{StringVal("C"), NumberVal(20)},
		{StringVal("D"), BoolVal(true)},
		{StringVal("E"), EmptyVal()},
	}

	tests := []dbTestCase{
		// --- basic counting ---
		{
			name: "count Apple profit values",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "count all profit values (blank criteria)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		{
			name: "count Pear yield values",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  2,
		},
		// --- field specified by column number ---
		{
			name: "field by column number 5 (Profit)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "field by column number 2 (Height)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    "2",
			wantType: ValueNumber,
			wantNum:  6,
		},
		// --- field specified by header string ---
		{
			name: "field header is case-insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- no matching records → 0 ---
		{
			name:     "no matching records returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- single record match → 1 ---
		{
			name: "single record match returns 1",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1,
		},
		// --- empty criteria (no criteria rows) matches all ---
		{
			name: "no criteria rows matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				// no condition rows
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		// --- multiple criteria rows (OR logic) ---
		{
			name: "OR criteria - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  5, // 3 Apple + 2 Pear
		},
		{
			name: "OR criteria - Apple OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  4, // 3 Apple + 1 Cherry
		},
		// --- multiple criteria columns (AND logic) ---
		{
			name: "AND criteria - Apple AND Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // Apple h=18 and h=14
		},
		{
			name: "three-column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1, // only Apple h=14, age=15
		},
		// --- combined AND/OR criteria ---
		{
			name: "combined AND/OR - (Apple AND Height>10) OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Cherry"), StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // Apple h>10: 2 records + Cherry: 1 record
		},
		{
			name: "combined AND/OR - (Apple AND Height>10) OR (Pear AND Height<10)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // Apple h>10: 2, Pear h<10: 1
		},
		// --- numeric comparison operators ---
		{
			name: "numeric comparison > on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  4, // h=18,12,13,14 all > 10
		},
		{
			name: "numeric comparison < on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // h=9 and h=8
		},
		{
			name: "numeric comparison >= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // h=18, h=14
		},
		{
			name: "numeric comparison <= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // h=12, h=9, h=8
		},
		{
			name: "numeric comparison <> on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  5, // all except Cherry (h=13)
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1, // only Apple h=14
		},
		// --- wildcard criteria ---
		{
			name: "wildcard * in criteria - trees starting with A",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // all Apple records
		},
		{
			name: "wildcard ? in criteria - Pea? matches Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // both Pear records
		},
		{
			name: "wildcard * contains pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // all trees contain 'e'
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		// --- counting a text column → 0 (DCOUNT only counts numbers) ---
		{
			name: "text column returns 0 - Tree column is text",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  0, // Tree column is text; DCOUNT ignores text
		},
		{
			name: "field index 1 on text column returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "1", // first column = Tree (text)
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- counting mixed types (only numbers counted) ---
		{
			name: "mixed types column - only numbers counted",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // only 10, 20
		},
		// --- field with empty cells (not counted) ---
		{
			name: "empty cells in field are not counted",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), NumberVal(200)},
				{StringVal("D"), EmptyVal()},
				{StringVal("E"), NumberVal(300)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  3, // 100, 200, 300 (empties not counted)
		},
		// --- cross-column criteria ---
		{
			name: "criteria on different column than counted field",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  2, // Age>14: age 20 and age 15
		},
		{
			name: "criteria on Profit, count Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  2, // Profit>100: Apple(105) and Cherry(105)
		},
		// --- database with boolean values ---
		{
			name: "boolean criteria match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C
		},
		{
			name: "boolean criteria match FALSE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(false)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  1, // only B
		},
		{
			name: "boolean column not counted by DCOUNT",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active")},
				{StringVal("A"), BoolVal(true)},
				{StringVal("B"), BoolVal(false)},
				{StringVal("C"), BoolVal(true)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Active"`,
			wantType: ValueNumber,
			wantNum:  0, // booleans are not numbers
		},
		// --- comparison with DCOUNTA behavior (DCOUNT only counts numbers) ---
		{
			name: "all text values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Note")},
				{StringVal("A"), StringVal("hello")},
				{StringVal("B"), StringVal("world")},
				{StringVal("C"), StringVal("test")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Note"`,
			wantType: ValueNumber,
			wantNum:  0, // DCOUNT skips text
		},
		{
			name: "all empty values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), EmptyVal()},
				{StringVal("B"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0, // empties are not numbers
		},
		// --- case-insensitive text criteria ---
		{
			name: "case-insensitive text criteria match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- criteria header not in database ---
		{
			name: "criteria header not in database - no match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(20)},
			},
			crit: [][]Value{
				{StringVal("NonExistentColumn")},
				{StringVal("A")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0, // criteria column not found -> never matches
		},
		// --- empty database ---
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				// no data rows
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- error field / field not found ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index negative returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- exact match with = prefix ---
		{
			name: "exact match with = prefix excludes partial matches",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Apple"), NumberVal(10)},
				{StringVal("Apple Pie"), NumberVal(20)},
				{StringVal("apple"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=Apple")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // "Apple" and "apple" match (case-insensitive), not "Apple Pie"
		},
		// --- <> on text ---
		{
			name: "not-equal text criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("<>Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // Pear(2) + Cherry(1)
		},
		// --- numeric criteria value (not comparison string) ---
		{
			name: "numeric criteria value exact match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score"), StringVal("Value")},
				{StringVal("A"), NumberVal(100), NumberVal(10)},
				{StringVal("B"), NumberVal(200), NumberVal(20)},
				{StringVal("C"), NumberVal(100), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Score")},
				{NumberVal(100)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C
		},
		// --- multiple OR rows with same column value ---
		{
			name: "multiple OR rows with exact height match",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=18")},
				{StringVal("=8")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // h=18 and h=8
		},
		// --- zeros count as numbers ---
		{
			name: "zeros are counted as numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(0)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  3, // zeros are numeric values
		},
		// --- negative numbers are counted ---
		{
			name: "negative numbers are counted",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(5)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- field as float column number ---
		{
			name: "field as float 2.9 truncates to column 2",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    "2.9",
			wantType: ValueNumber,
			wantNum:  6, // column 2 = Height, all numeric
		},
		// --- criteria on same column as field ---
		{
			name: "criteria on same column as field",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  3, // h=18, h=13, h=14
		},
		// --- = empty matches only empty cells ---
		{
			name: "= empty criteria matches empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(20)},
				{StringVal("C"), NumberVal(30)},
				{EmptyVal(), NumberVal(40)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  2, // rows with empty Name
		},
		// --- <> empty matches non-empty cells ---
		{
			name: "<> empty criteria matches non-empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(20)},
				{StringVal("C"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("<>")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C (non-empty Name)
		},
		// --- whitespace-padded criteria header ---
		{
			name: "whitespace-padded criteria header still matches",
			db:   db,
			crit: [][]Value{
				{StringVal(" Tree ")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- numeric string in field column not counted ---
		{
			name: "numeric string values not counted by DCOUNT",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("100")},
				{StringVal("B"), StringVal("200")},
				{StringVal("C"), NumberVal(300)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1, // only C is a true number
		},
		// --- duplicate criteria rows count once ---
		{
			name: "duplicate criteria rows do not double-count",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1, // only one Cherry record; OR of same criteria doesn't duplicate
		},
		// --- multiple AND columns with OR rows ---
		{
			name: "multiple AND columns with multiple OR rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Yield")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<9")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // Apple yield>10: h=18(y=14); Pear yield<9: h=9(y=8)
		},
		// --- single cell database (headers only, no data) ---
		{
			name: "single column database with no matching data",
			db: [][]Value{
				{StringVal("X")},
			},
			crit: [][]Value{
				{StringVal("X")},
				{StringVal("")},
			},
			field:    `"X"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- mixed positive and negative with criteria ---
		{
			name: "count mixed positive and negative matching criteria",
			db: [][]Value{
				{StringVal("Type"), StringVal("Value")},
				{StringVal("A"), NumberVal(-5)},
				{StringVal("B"), NumberVal(10)},
				{StringVal("A"), NumberVal(-3)},
				{StringVal("B"), NumberVal(7)},
				{StringVal("A"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Type")},
				{StringVal("A")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  3, // -5, -3, 0 all numeric
		},
		// --- very large values ---
		{
			name: "very large numbers are counted",
			db: [][]Value{
				{StringVal("ID"), StringVal("Amount")},
				{StringVal("A"), NumberVal(1e15)},
				{StringVal("B"), NumberVal(1e-15)},
			},
			crit: [][]Value{
				{StringVal("ID")},
				{StringVal("")},
			},
			field:    `"Amount"`,
			wantType: ValueNumber,
			wantNum:  2,
		},
		// --- criteria with string comparison operators on text ---
		{
			name: "string comparison > on text column",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Alpha"), NumberVal(1)},
				{StringVal("Beta"), NumberVal(2)},
				{StringVal("Gamma"), NumberVal(3)},
				{StringVal("Delta"), NumberVal(4)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal(">Delta")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1, // only Gamma > Delta lexically
		},
	}

	runDBTests(t, "DCOUNT", tests)
}

func TestDCOUNT_WrongArgCount(t *testing.T) {
	// Direct function call with wrong arg counts
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDCount(tt.args)
			if err != nil {
				t.Fatalf("fnDCount error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDCount(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDCOUNT_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the counted field, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DCOUNT(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DCOUNT with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDCOUNT_ErrorInDatabase(t *testing.T) {
	// Error in a non-counted field should not affect counting.
	db := [][]Value{
		{StringVal("Name"), StringVal("Score"), StringVal("Value")},
		{StringVal("A"), NumberVal(100), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValNA), NumberVal(20)},
		{StringVal("C"), NumberVal(300), NumberVal(30)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DCOUNT(A1:C4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Error is in "Score" column but we're counting "Value" column, so no error.
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("DCOUNT = %+v, want 3", got)
	}
}

// ---------------------------------------------------------------------------
// DCOUNTA
// ---------------------------------------------------------------------------

func TestDCOUNTA(t *testing.T) {
	db := standardDB()

	mixedDB := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), StringVal("text")},
		{StringVal("C"), NumberVal(20)},
		{StringVal("D"), BoolVal(true)},
		{StringVal("E"), EmptyVal()},
	}

	tests := []dbTestCase{
		// --- basic counting ---
		{
			name: "basic single text criteria - Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "count all profit values (blank criteria)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		{
			name: "count Pear yield values",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  2,
		},
		// --- field specified by column number ---
		{
			name: "field by column number 5 (Profit)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "field by column number 1 (Tree - text column)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "1",
			wantType: ValueNumber,
			wantNum:  3, // DCOUNTA counts text, unlike DCOUNT
		},
		// --- field specified by header string ---
		{
			name: "field header is case-insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- no matching records → 0 ---
		{
			name:     "no matching records returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- all records match (empty criteria) ---
		{
			name: "no criteria rows matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				// no condition rows
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		// --- single record match → 1 ---
		{
			name: "single record match returns 1",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1,
		},
		// --- multiple criteria rows (OR logic) ---
		{
			name: "OR criteria - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  5, // 3 Apple + 2 Pear
		},
		{
			name: "OR criteria - Apple OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  4, // 3 Apple + 1 Cherry
		},
		// --- multiple criteria columns (AND logic) ---
		{
			name: "AND criteria - Apple AND Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // Apple h=18 and h=14
		},
		{
			name: "three-column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1, // only Apple h=14, age=15
		},
		// --- combined AND/OR criteria ---
		{
			name: "combined AND/OR - (Apple AND Height>10) OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Cherry"), StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // Apple h>10: 2 records + Cherry: 1 record
		},
		// --- numeric comparison operators ---
		{
			name: "numeric comparison > on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  4, // h=18,12,13,14 all > 10
		},
		{
			name: "numeric comparison < on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // h=9 and h=8
		},
		{
			name: "numeric comparison >= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // h=18, h=14
		},
		{
			name: "numeric comparison <= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // h=12, h=9, h=8
		},
		{
			name: "numeric comparison <> on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  5, // all except Cherry (h=13)
		},
		// --- wildcard criteria ---
		{
			name: "wildcard * in criteria - trees starting with A",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // all Apple records
		},
		{
			name: "wildcard ? in criteria - Pea? matches Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2, // both Pear records
		},
		{
			name: "wildcard * contains pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // all trees contain 'e'
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		// --- DCOUNTA counts text values (unlike DCOUNT) ---
		{
			name: "text column counted by DCOUNTA - Tree column",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  6, // DCOUNTA counts text; DCOUNT would return 0
		},
		{
			name: "all text values counted by DCOUNTA",
			db: [][]Value{
				{StringVal("Name"), StringVal("Note")},
				{StringVal("A"), StringVal("hello")},
				{StringVal("B"), StringVal("world")},
				{StringVal("C"), StringVal("test")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Note"`,
			wantType: ValueNumber,
			wantNum:  3, // DCOUNTA counts text; DCOUNT returns 0
		},
		// --- DCOUNTA counts boolean values ---
		{
			name: "boolean column counted by DCOUNTA",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active")},
				{StringVal("A"), BoolVal(true)},
				{StringVal("B"), BoolVal(false)},
				{StringVal("C"), BoolVal(true)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Active"`,
			wantType: ValueNumber,
			wantNum:  3, // DCOUNTA counts booleans; DCOUNT returns 0
		},
		// --- DCOUNTA counts numbers ---
		{
			name: "numeric column counted by DCOUNTA",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		{
			name: "zeros counted by DCOUNTA",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(0)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  3, // zeros are non-empty
		},
		// --- DCOUNTA does NOT count empty cells ---
		{
			name: "empty cells not counted by DCOUNTA",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), NumberVal(200)},
				{StringVal("D"), EmptyVal()},
				{StringVal("E"), NumberVal(300)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  3, // 100, 200, 300 (empties not counted)
		},
		{
			name: "all empty values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), EmptyVal()},
				{StringVal("B"), EmptyVal()},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- mixed types (all counted except empty) ---
		{
			name: "mixed types - count non-empty",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  4, // 10, "text", 20, true (not empty E)
		},
		// --- cross-column criteria ---
		{
			name: "criteria on different column than counted field",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  2, // Age>14: age 20 and age 15
		},
		{
			name: "criteria on Profit, count Tree (text)",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  2, // Profit>100: Apple(105) and Cherry(105) — Tree is text, DCOUNTA counts it
		},
		// --- case-insensitive text criteria ---
		{
			name: "case-insensitive text criteria match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- exact match with = prefix ---
		{
			name: "exact match with = prefix excludes partial matches",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Apple"), NumberVal(10)},
				{StringVal("Apple Pie"), NumberVal(20)},
				{StringVal("apple"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=Apple")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // "Apple" and "apple" match (case-insensitive), not "Apple Pie"
		},
		// --- <> on text ---
		{
			name: "not-equal text criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("<>Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3, // Pear(2) + Cherry(1)
		},
		// --- criteria header not in database ---
		{
			name: "criteria header not in database - no match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(20)},
			},
			crit: [][]Value{
				{StringVal("NonExistentColumn")},
				{StringVal("A")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0, // criteria column not found -> never matches
		},
		// --- empty database ---
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				// no data rows
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- error field / field not found ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index negative returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- numeric criteria value (not comparison string) ---
		{
			name: "numeric criteria value exact match",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score"), StringVal("Value")},
				{StringVal("A"), NumberVal(100), StringVal("x")},
				{StringVal("B"), NumberVal(200), StringVal("y")},
				{StringVal("C"), NumberVal(100), StringVal("z")},
			},
			crit: [][]Value{
				{StringVal("Score")},
				{NumberVal(100)},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C
		},
		// --- boolean criteria ---
		{
			name: "boolean criteria match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C
		},
		{
			name: "boolean criteria match FALSE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(false)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  1, // only B
		},
		// --- DCOUNTA counts numeric strings ---
		{
			name: "numeric strings are counted by DCOUNTA",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("100")},
				{StringVal("B"), StringVal("200")},
				{StringVal("C"), NumberVal(300)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  3, // all non-empty: "100", "200", 300
		},
		// --- field as float column number ---
		{
			name: "field as float 1.7 truncates to column 1",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "1.7",
			wantType: ValueNumber,
			wantNum:  3, // column 1 = Tree (text), DCOUNTA counts text
		},
		// --- criteria on same column as field ---
		{
			name: "criteria on same column as field",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  3, // h=18, h=13, h=14
		},
		// --- = empty matches only empty cells ---
		{
			name: "= empty criteria matches empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("x")},
				{EmptyVal(), StringVal("y")},
				{StringVal("C"), StringVal("z")},
				{EmptyVal(), StringVal("w")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // rows with empty Name
		},
		// --- <> empty matches non-empty cells ---
		{
			name: "<> empty criteria matches non-empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("x")},
				{EmptyVal(), StringVal("y")},
				{StringVal("C"), StringVal("z")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("<>")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // A and C
		},
		// --- whitespace-padded criteria header ---
		{
			name: "whitespace-padded criteria header still matches",
			db:   db,
			crit: [][]Value{
				{StringVal(" Tree ")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		// --- multiple empty cells among other types ---
		{
			name: "multiple empty cells among mixed types",
			db: [][]Value{
				{StringVal("Name"), StringVal("Data")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), StringVal("hi")},
				{StringVal("D"), EmptyVal()},
				{StringVal("E"), BoolVal(false)},
				{StringVal("F"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Data"`,
			wantType: ValueNumber,
			wantNum:  3, // 1, "hi", false (3 empties skipped)
		},
		// --- duplicate criteria rows ---
		{
			name: "duplicate criteria rows do not double count",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  1, // only one Cherry record
		},
		// --- multiple AND columns with OR rows ---
		{
			name: "multiple AND columns with multiple OR rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Yield")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<9")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  2, // Apple y>10: Apple(y=14); Pear y<9: Pear(y=8) -> count text
		},
		// --- single cell database ---
		{
			name: "single column headers-only database returns 0",
			db: [][]Value{
				{StringVal("X")},
			},
			crit: [][]Value{
				{StringVal("X")},
				{StringVal("")},
			},
			field:    `"X"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- DCOUNTA counts everything except empty in a fully-populated column ---
		{
			name: "fully-populated mixed column counts all non-empty",
			db: [][]Value{
				{StringVal("Name"), StringVal("Stuff")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), StringVal("")},
				{StringVal("C"), BoolVal(true)},
				{StringVal("D"), NumberVal(-99)},
				{StringVal("E"), StringVal("hello")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Stuff"`,
			wantType: ValueNumber,
			wantNum:  5, // 0, "", true, -99, "hello" all non-empty
		},
		// --- large number of matching records ---
		{
			name: "count with string comparison > on text",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Alpha"), StringVal("x")},
				{StringVal("Beta"), StringVal("y")},
				{StringVal("Gamma"), StringVal("z")},
				{StringVal("Delta"), StringVal("w")},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal(">Delta")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1, // Gamma > Delta lexically
		},
	}

	runDBTests(t, "DCOUNTA", tests)
}

func TestDCOUNTA_WrongArgCount(t *testing.T) {
	// Direct function call with wrong arg counts
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDCountA(tt.args)
			if err != nil {
				t.Fatalf("fnDCountA error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDCountA(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDCOUNTA_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the counted field, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DCOUNTA(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DCOUNTA with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDCOUNTA_ErrorInNonCountedField(t *testing.T) {
	// Error in a non-counted field should not affect counting.
	db := [][]Value{
		{StringVal("Name"), StringVal("Score"), StringVal("Value")},
		{StringVal("A"), NumberVal(100), StringVal("x")},
		{StringVal("B"), ErrorVal(ErrValNA), StringVal("y")},
		{StringVal("C"), NumberVal(300), StringVal("z")},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DCOUNTA(A1:C4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Error is in "Score" column but we're counting "Value" column, so no error.
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("DCOUNTA = %+v, want 3", got)
	}
}

// ---------------------------------------------------------------------------
// DGET
// ---------------------------------------------------------------------------

func TestDGET(t *testing.T) {
	db := standardDB()

	// Database with mixed value types for targeted tests.
	mixedDB := [][]Value{
		{StringVal("Name"), StringVal("Value"), StringVal("Active")},
		{StringVal("Alpha"), NumberVal(10), BoolVal(true)},
		{StringVal("Beta"), StringVal("hello"), BoolVal(false)},
		{StringVal("Gamma"), NumberVal(30), BoolVal(true)},
		{StringVal("Delta"), EmptyVal(), BoolVal(false)},
		{StringVal("Epsilon"), NumberVal(50), BoolVal(true)},
	}

	// Single-record database.
	singleDB := [][]Value{
		{StringVal("Item"), StringVal("Price"), StringVal("Qty")},
		{StringVal("Widget"), NumberVal(9.99), NumberVal(42)},
	}

	tests := []dbTestCase{
		// --- basic: single matching record returns the field value ---
		{
			name: "single match returns numeric value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- no matching records → #VALUE! ---
		{
			name:     "no matches returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- multiple matching records → #NUM! ---
		{
			name:     "multiple matches returns NUM error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- retrieving string values ---
		{
			name: "returns string value from matching record",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Cherry"), StringVal("=13")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Cherry",
		},
		// --- AND criteria (multiple columns in one criteria row) ---
		{
			name: "AND criteria narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal("=18")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- AND criteria narrowing Apple by Age ---
		{
			name: "AND criteria Apple with Age=15",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal("=15")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75,
		},
		// --- field specified by column number ---
		{
			name: "field by column number 5 (Profit)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "field by column number 1 (Tree)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Cherry"), StringVal("=13")},
			},
			field:    "1",
			wantType: ValueString,
			wantStr:  "Cherry",
		},
		// --- case-insensitive field matching ---
		{
			name: "field name case insensitive - lowercase",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "field name case insensitive - mixed case",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"PROFIT"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- numeric comparison operators ---
		{
			name: "greater than narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">17")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Apple", // only Apple has Height=18 > 17
		},
		{
			name: "less than narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal("<9")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Pear", // Pear age=8
		},
		{
			name: "greater-equal narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=18")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "less-equal narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=8")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Apple", // Apple row with Height=8
		},
		{
			name: "not-equal produces multiple matches → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("<>Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM, // Cherry + 2 Pears = 3 matches
		},
		// --- OR criteria (multiple criteria rows) → multiple matches → #NUM! ---
		{
			name: "OR criteria causes multiple matches returns NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM, // Cherry + 2 Pears = 3 matches
		},
		// --- OR criteria where each row matches one record → still #NUM! ---
		{
			name: "OR criteria two rows each matching one record → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal("=18")},
				{StringVal("Cherry"), StringVal("=13")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- wildcard * ---
		{
			name: "wildcard * matches single unique prefix",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Ch*")}, // matches Cherry only
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "wildcard * matches multiple → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")}, // matches all 3 Apples
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- wildcard ? ---
		{
			name: "wildcard ? matches multiple → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")}, // matches Pear (x2)
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		{
			name: "wildcard ? narrows to single match with AND",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Pea?"), StringVal("=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96, // only Pear with Age=12
		},
		// --- exact match with = prefix ---
		{
			name: "exact match with = prefix",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("=Cherry")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13,
		},
		// --- cross-column criteria: use different columns in criteria vs field ---
		{
			name: "cross-column criteria - get Age from Height criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=9")},
			},
			field:    `"Age"`,
			wantType: ValueNumber,
			wantNum:  8, // Pear has Height=9, Age=8
		},
		// --- retrieving boolean values ---
		{
			name: "retrieve boolean value true",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Alpha")},
			},
			field:    `"Active"`,
			wantType: ValueBool,
			wantBool: true,
		},
		{
			name: "retrieve boolean value false",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Beta")},
			},
			field:    `"Active"`,
			wantType: ValueBool,
			wantBool: false,
		},
		// --- retrieving text values from mixed DB ---
		{
			name: "retrieve text value from mixed DB",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Beta")},
			},
			field:    `"Value"`,
			wantType: ValueString,
			wantStr:  "hello",
		},
		// --- empty cell in matching record ---
		{
			name: "retrieve empty cell value from matching record",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Delta")},
			},
			field:    `"Value"`,
			wantType: ValueEmpty,
		},
		// --- single record database: always a single match ---
		{
			name: "single record DB returns the value",
			db:   singleDB,
			crit: [][]Value{
				{StringVal("Item")},
				{StringVal("Widget")},
			},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  9.99,
		},
		{
			name: "single record DB blank criteria returns the value",
			db:   singleDB,
			crit: [][]Value{
				{StringVal("Item")},
				{StringVal("")},
			},
			field:    `"Qty"`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		// --- all records match empty criteria → #NUM! (if >1 record) ---
		{
			name: "blank criteria on multi-record DB returns NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		{
			name: "no criteria rows matches all → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- field name not found → #VALUE! ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- field index out of range → #VALUE! ---
		{
			name:     "field index too large returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index zero returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index negative returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- criteria matching single unique record among many ---
		{
			name: "unique Yield value identifies single record",
			db:   db,
			crit: [][]Value{
				{StringVal("Yield")},
				{StringVal("=14")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Apple", // only Apple row 1 has Yield=14
		},
		{
			name: "unique Profit=76.8 identifies single Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal("=76.8")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Pear",
		},
		// --- AND criteria with numeric comparisons narrowing to single match ---
		{
			name: "AND with > and < narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Height"), StringVal("Age")},
				{StringVal(">13"), StringVal("<16")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Apple", // Apple Height=14, Age=15
		},
		// --- case-insensitive criteria value matching ---
		{
			name: "criteria value is case insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("cherry")}, // lowercase criteria for "Cherry"
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "criteria value uppercase matches",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("CHERRY")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13,
		},
		// --- case-insensitive criteria header matching ---
		{
			name: "criteria header is case insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("tree")}, // lowercase header
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- error in database cell propagates ---
		{
			name: "error cell in matching record propagates",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("Alice"), NumberVal(100)},
				{StringVal("Bob"), ErrorVal(ErrValNA)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("Bob")},
			},
			field:    `"Score"`,
			wantType: ValueError,
			wantErr:  ErrValNA,
		},
		// --- not-equal zero operator ---
		{
			name: "not-equal zero criteria <>0",
			db: [][]Value{
				{StringVal("Item"), StringVal("Qty")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{
				{StringVal("Qty")},
				{StringVal("<>0")},
			},
			field:    `"Item"`,
			wantType: ValueString,
			wantStr:  "B", // only B has Qty != 0
		},
		// --- wildcard contains pattern *text* ---
		{
			name: "wildcard contains pattern *ear* matches single",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("*ear*"), StringVal("=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96, // Pear with Age=12
		},
		{
			name: "wildcard contains pattern *ppl* matches multiple → NUM error",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*ppl*")}, // matches all 3 Apples
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- criteria with not-equal on text using <> ---
		{
			name: "not-equal text criteria with AND narrows to single",
			db: [][]Value{
				{StringVal("Color"), StringVal("Size")},
				{StringVal("Red"), NumberVal(5)},
				{StringVal("Blue"), NumberVal(10)},
			},
			crit: [][]Value{
				{StringVal("Color")},
				{StringVal("<>Red")},
			},
			field:    `"Size"`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		// --- numeric criteria as actual number (not string) ---
		{
			name: "numeric criteria value matches exact number",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{NumberVal(13)}, // numeric 13 matches Cherry height=13
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Cherry",
		},
		// --- boolean criteria in criteria range ---
		{
			name: "boolean TRUE criteria matches boolean field",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Active"), StringVal("Name")},
				{BoolVal(true), StringVal("Gamma")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  30,
		},
		{
			name: "boolean FALSE criteria matches boolean field",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Active"), StringVal("Name")},
				{BoolVal(false), StringVal("Beta")},
			},
			field:    `"Value"`,
			wantType: ValueString,
			wantStr:  "hello",
		},
		// --- criteria header not in database → no match → VALUE error ---
		{
			name: "criteria header not in DB causes no match → VALUE error",
			db:   db,
			crit: [][]Value{
				{StringVal("Color")}, // "Color" not in DB
				{StringVal("Red")},
			},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE, // no records match → VALUE
		},
		// --- mixed types: string cell vs numeric criteria ---
		{
			name: "numeric criteria does not match string cell",
			db: [][]Value{
				{StringVal("Label"), StringVal("Val")},
				{StringVal("X"), StringVal("hello")},
				{StringVal("Y"), NumberVal(42)},
			},
			crit: [][]Value{
				{StringVal("Val")},
				{NumberVal(42)},
			},
			field:    `"Label"`,
			wantType: ValueString,
			wantStr:  "Y", // only Y has numeric 42
		},
		// --- field by column number 2 (middle column) ---
		{
			name: "field by column number 2 (Height)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    "2",
			wantType: ValueNumber,
			wantNum:  13,
		},
		// --- three-column AND criteria narrowing to single match ---
		{
			name: "three-column AND criteria narrows to one record",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<16")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  10, // Apple, Height=14, Age=15 → Yield=10
		},
		// --- single record DB with no criteria rows → returns the single record ---
		{
			name: "single record DB no criteria rows returns value",
			db:   singleDB,
			crit: [][]Value{
				{StringVal("Item")},
			},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  9.99,
		},
		// --- decimal precision in matching ---
		{
			name: "decimal criteria exact match",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{NumberVal(76.8)},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Pear",
		},
		// --- less-equal string comparison operator ---
		{
			name: "less-equal on age narrows to single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal("<=9")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // Apple with Age=9, Profit=45
		},
		// --- wildcard with = prefix ---
		{
			name: "wildcard with = prefix matches pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("=Ch*")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13, // Cherry
		},
		// --- field index last column ---
		{
			name: "field index equals number of columns",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    "5", // last column (Profit)
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- database with all identical values in field, single match by other col ---
		{
			name: "identical field values returns matched record value",
			db: [][]Value{
				{StringVal("ID"), StringVal("Score")},
				{NumberVal(1), NumberVal(100)},
				{NumberVal(2), NumberVal(100)},
				{NumberVal(3), NumberVal(100)},
			},
			crit: [][]Value{
				{StringVal("ID")},
				{NumberVal(2)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  100,
		},
		// --- empty value in criteria with non-empty header still matches all ---
		{
			name: "empty criteria with EmptyVal matches all single-row DB",
			db:   singleDB,
			crit: [][]Value{
				{StringVal("Item")},
				{EmptyVal()},
			},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  9.99,
		},
	}

	runDBTests(t, "DGET", tests)
}

func TestDGET_WrongArgCount(t *testing.T) {
	result, err := fnDGet(nil)
	if err != nil {
		t.Fatalf("fnDGet error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDGet(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DMAX
// ---------------------------------------------------------------------------

func TestDMAX(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- basic: max of numeric column with single text criteria ---
		{
			name: "max Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "max Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96,
		},
		{
			name: "max Cherry profit (single record)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "max Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		// --- all records match (blank criteria) ---
		{
			name: "max all height - blank criteria value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  18,
		},
		{
			name: "max all profit - blank criteria value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "max all records - no criteria rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Age"`,
			wantType: ValueNumber,
			wantNum:  20,
		},
		// --- no matching records → 0 ---
		{
			name:     "no matching records returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- single record match → that value ---
		{
			name: "single record match returns that value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13,
		},
		// --- field specified by column number ---
		{
			name: "field by column number 5 (Profit)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "field by column number 2 (Height)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    "2",
			wantType: ValueNumber,
			wantNum:  18,
		},
		// --- field specified by header string (case-insensitive) ---
		{
			name: "field header is case-insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- multiple criteria rows (OR logic) ---
		{
			name: "OR criteria - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // max of Apple(105,75,45) and Pear(96,76.8)
		},
		{
			name: "OR criteria - Apple OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  18, // max of Apple(18,14,8) and Cherry(13)
		},
		// --- multiple criteria columns (AND logic) ---
		{
			name: "AND criteria - Apple AND Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // Apple h=18(105), h=14(75)
		},
		{
			name: "three-column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14, age=15
		},
		// --- combined AND/OR criteria ---
		{
			name: "combined AND/OR - (Apple AND Height>10) OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Cherry"), StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // Apple h>10: 105,75; Cherry: 105
		},
		{
			name: "combined AND/OR - (Apple AND Height>10) OR (Pear AND Height<10)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // Apple h>10: 105,75; Pear h<10: 76.8
		},
		// --- numeric comparison operators ---
		{
			name: "numeric comparison > on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  14, // h=18(y=14), h=12(y=10), h=13(y=9), h=14(y=10)
		},
		{
			name: "numeric comparison < on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  76.8, // h=9(76.8), h=8(45)
		},
		{
			name: "numeric comparison >= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // h=18(105), h=14(75)
		},
		{
			name: "numeric comparison <= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96, // h=12(96), h=9(76.8), h=8(45)
		},
		{
			name: "numeric comparison <> on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // all except Cherry(h=13)
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14
		},
		// --- wildcard criteria ---
		{
			name: "wildcard * in criteria - trees starting with A",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105, // all Apple records: 105, 75, 45
		},
		{
			name: "wildcard ? in criteria - Pea? matches Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  12, // Pear: h=12, h=9
		},
		{
			name: "wildcard * contains pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // Apple, Pear, Cherry all contain 'e'
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- cross-column criteria ---
		{
			name: "criteria on Age, max of Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  14, // age>14: age=20(y=14), age=15(y=10)
		},
		{
			name: "criteria on Profit, max of Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  18, // Profit>100: Apple(h=18,p=105), Cherry(h=13,p=105)
		},
		// --- negative numbers (max of negatives) ---
		{
			name: "max with all negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(-20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -5,
		},
		{
			name: "max negatives with criteria",
			db: [][]Value{
				{StringVal("Cat"), StringVal("Score")},
				{StringVal("X"), NumberVal(-100)},
				{StringVal("Y"), NumberVal(-3)},
				{StringVal("X"), NumberVal(-50)},
			},
			crit:     [][]Value{{StringVal("Cat")}, {StringVal("X")}},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  -50, // max of -100 and -50
		},
		// --- all same values → that value ---
		{
			name: "all same values returns that value",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
				{StringVal("C"), NumberVal(42)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		// --- decimal values ---
		{
			name: "max of decimal values",
			db: [][]Value{
				{StringVal("Item"), StringVal("Price")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.75)},
				{StringVal("C"), NumberVal(0.99)},
				{StringVal("D"), NumberVal(2.74)},
			},
			crit:     [][]Value{{StringVal("Item")}, {StringVal("")}},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  2.75,
		},
		{
			name: "max Pear profit is decimal",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96, // Pear: 96, 76.8
		},
		// --- text column → 0 (DMAX only considers numbers) ---
		{
			name: "text column returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- database with empty cells in field ---
		{
			name: "empty cells in field are ignored",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), NumberVal(200)},
				{StringVal("D"), EmptyVal()},
				{StringVal("E"), NumberVal(50)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  200,
		},
		// --- mixed types (only numbers considered) ---
		{
			name: "mixed types - only numbers considered for max",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  20,
		},
		// --- boolean criteria match ---
		{
			name: "boolean criteria match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  30, // A=10, C=30
		},
		{
			name: "boolean criteria match FALSE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(false)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  20, // only B
		},
		// --- large numbers ---
		{
			name: "max with large numbers",
			db: [][]Value{
				{StringVal("ID"), StringVal("Amount")},
				{StringVal("A"), NumberVal(1e10)},
				{StringVal("B"), NumberVal(1e12)},
				{StringVal("C"), NumberVal(1e8)},
			},
			crit:     [][]Value{{StringVal("ID")}, {StringVal("")}},
			field:    `"Amount"`,
			wantType: ValueNumber,
			wantNum:  1e12,
		},
		// --- zero values ---
		{
			name: "max when values include zero",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(0)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- single row database ---
		{
			name: "single row database returns that value",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Only"), NumberVal(99)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  99,
		},
		// --- mixed positive and negative ---
		{
			name: "max of mixed positive and negative",
			db: [][]Value{
				{StringVal("Type"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("A"), NumberVal(5)},
				{StringVal("A"), NumberVal(-3)},
				{StringVal("A"), NumberVal(2)},
			},
			crit:     [][]Value{{StringVal("Type")}, {StringVal("A")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  5,
		},
		// --- very close values (precision) ---
		{
			name: "max distinguishes very close values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.0000001)},
				{StringVal("B"), NumberVal(1.0000002)},
				{StringVal("C"), NumberVal(1.0000000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1.0000002,
		},
		// --- field as float column number ---
		{
			name: "field as float 5.5 truncates to column 5",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5.5",
			wantType: ValueNumber,
			wantNum:  105, // column 5 = Profit
		},
		// --- criteria on same column as field ---
		{
			name: "criteria on same column as field - max of Height where Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  18,
		},
		// --- = empty matches only empty cells ---
		{
			name: "= empty criteria matches empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(50)},
				{StringVal("C"), NumberVal(30)},
				{EmptyVal(), NumberVal(40)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  50, // max of 50 and 40
		},
		// --- <> empty matches non-empty cells ---
		{
			name: "<> empty criteria matches non-empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(50)},
				{StringVal("C"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("<>")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  30, // max of 10 and 30 (empty Name row excluded)
		},
		// --- whitespace-padded criteria header ---
		{
			name: "whitespace-padded criteria header still matches",
			db:   db,
			crit: [][]Value{
				{StringVal(" Tree ")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		// --- empty database ---
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- duplicate criteria rows ---
		{
			name: "duplicate criteria rows still max correctly",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96,
		},
		// --- multiple AND columns with OR rows ---
		{
			name: "multiple AND columns with multiple OR rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Yield")},
				{StringVal("Apple"), StringVal("<8")},
				{StringVal("Pear"), StringVal(">9")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  96, // Apple y<8: Apple(y=6,p=45); Pear y>9: Pear(y=10,p=96) -> max=96
		},
		// --- field not found ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- field index errors ---
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index negative returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- single negative value ---
		{
			name: "single negative matching value",
			db: [][]Value{
				{StringVal("Cat"), StringVal("Val")},
				{StringVal("X"), NumberVal(-42)},
				{StringVal("Y"), NumberVal(10)},
			},
			crit:     [][]Value{{StringVal("Cat")}, {StringVal("X")}},
			field:    `"Val"`,
			wantType: ValueNumber,
			wantNum:  -42,
		},
		// --- boolean column not considered for max ---
		{
			name: "boolean column returns 0 for max",
			db: [][]Value{
				{StringVal("Name"), StringVal("Flag")},
				{StringVal("A"), BoolVal(true)},
				{StringVal("B"), BoolVal(false)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Flag"`,
			wantType: ValueNumber,
			wantNum:  0, // booleans not numeric
		},
		// --- case-insensitive criteria ---
		{
			name: "case-insensitive text criteria match for max",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("APPLE")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
	}

	runDBTests(t, "DMAX", tests)
}

func TestDMAX_WrongArgCount(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDMax(tt.args)
			if err != nil {
				t.Fatalf("fnDMax error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDMax(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDMAX_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the field column, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DMAX(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DMAX with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDMAX_ErrorInNonFieldColumn(t *testing.T) {
	// Error in a non-field column should not affect the result.
	db := [][]Value{
		{StringVal("Name"), StringVal("Score"), StringVal("Value")},
		{StringVal("A"), NumberVal(100), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValNA), NumberVal(20)},
		{StringVal("C"), NumberVal(300), NumberVal(30)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DMAX(A1:C4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("DMAX = %+v, want 30", got)
	}
}

// ---------------------------------------------------------------------------
// DMIN
// ---------------------------------------------------------------------------

func TestDMIN(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- basic: min of numeric column with single text criteria ---
		{
			name: "min Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		{
			name: "min Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  76.8,
		},
		{
			name: "min Cherry profit (single record)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "min Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  8,
		},
		// --- all records match (blank criteria) ---
		{
			name: "min all height - blank criteria value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  8,
		},
		{
			name: "min all profit - blank criteria value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		{
			name: "min all records - no criteria rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Age"`,
			wantType: ValueNumber,
			wantNum:  8,
		},
		// --- no matching records → 0 ---
		{
			name:     "no matching records returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- single record match → that value ---
		{
			name: "single record match returns that value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13,
		},
		// --- field specified by column number ---
		{
			name: "field by column number 5 (Profit)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  45,
		},
		{
			name: "field by column number 2 (Height)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    "2",
			wantType: ValueNumber,
			wantNum:  8,
		},
		// --- field specified by header string (case-insensitive) ---
		{
			name: "field header is case-insensitive",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		// --- multiple criteria rows (OR logic) ---
		{
			name: "OR criteria - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // min of Apple(105,75,45) and Pear(96,76.8)
		},
		{
			name: "OR criteria - Apple OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  8, // min of Apple(18,14,8) and Cherry(13)
		},
		// --- multiple criteria columns (AND logic) ---
		{
			name: "AND criteria - Apple AND Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // Apple h=18(105), h=14(75)
		},
		{
			name: "three-column AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<20")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14, age=15
		},
		// --- combined AND/OR criteria ---
		{
			name: "combined AND/OR - (Apple AND Height>10) OR Cherry",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Cherry"), StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // Apple h>10: 105,75; Cherry: 105 → min=75
		},
		{
			name: "combined AND/OR - (Apple AND Height>10) OR (Pear AND Height<10)",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // Apple h>10: 105,75; Pear h<10: 76.8 → min=75
		},
		// --- numeric comparison operators ---
		{
			name: "numeric comparison > on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  9, // h=18(y=14), h=12(y=10), h=13(y=9), h=14(y=10) → min=9
		},
		{
			name: "numeric comparison < on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // h=9(76.8), h=8(45) → min=45
		},
		{
			name: "numeric comparison >= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // h=18(105), h=14(75) → min=75
		},
		{
			name: "numeric comparison <= on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // h=12(96), h=9(76.8), h=8(45) → min=45
		},
		{
			name: "numeric comparison <> on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // all except Cherry(h=13) → min=45
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14
		},
		// --- wildcard criteria ---
		{
			name: "wildcard * in criteria - trees starting with A",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // all Apple records: 105, 75, 45 → min=45
		},
		{
			name: "wildcard ? in criteria - Pea? matches Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  9, // Pear: h=12, h=9 → min=9
		},
		{
			name: "wildcard * contains pattern",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("*e*")}, // Apple, Pear, Cherry all contain 'e'
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // all records → min=45
		},
		// --- cross-column criteria ---
		{
			name: "criteria on Age, min of Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Age")},
				{StringVal(">14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  10, // age>14: age=20(y=14), age=15(y=10) → min=10
		},
		{
			name: "criteria on Profit, min of Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Profit")},
				{StringVal(">100")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  13, // Profit>100: Apple(h=18,p=105), Cherry(h=13,p=105) → min=13
		},
		// --- negative numbers (min of negatives) ---
		{
			name: "min with all negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(-20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -20,
		},
		{
			name: "min negatives with criteria",
			db: [][]Value{
				{StringVal("Cat"), StringVal("Score")},
				{StringVal("X"), NumberVal(-100)},
				{StringVal("Y"), NumberVal(-3)},
				{StringVal("X"), NumberVal(-50)},
			},
			crit:     [][]Value{{StringVal("Cat")}, {StringVal("X")}},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  -100, // min of -100 and -50
		},
		// --- all same values → that value ---
		{
			name: "all same values returns that value",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
				{StringVal("C"), NumberVal(42)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		// --- decimal values ---
		{
			name: "min of decimal values",
			db: [][]Value{
				{StringVal("Item"), StringVal("Price")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.75)},
				{StringVal("C"), NumberVal(0.99)},
				{StringVal("D"), NumberVal(2.74)},
			},
			crit:     [][]Value{{StringVal("Item")}, {StringVal("")}},
			field:    `"Price"`,
			wantType: ValueNumber,
			wantNum:  0.99,
		},
		{
			name: "min Pear profit is decimal",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  76.8, // Pear: 96, 76.8
		},
		// --- text column → 0 (DMIN only considers numbers) ---
		{
			name: "text column returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- database with empty cells in field ---
		{
			name: "empty cells in field are ignored",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), NumberVal(200)},
				{StringVal("D"), EmptyVal()},
				{StringVal("E"), NumberVal(50)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  50,
		},
		// --- mixed types (only numbers considered) ---
		{
			name: "mixed types - only numbers considered for min",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		// --- boolean criteria match ---
		{
			name: "boolean criteria match TRUE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(true)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  10, // A=10, C=30 → min=10
		},
		{
			name: "boolean criteria match FALSE",
			db: [][]Value{
				{StringVal("Name"), StringVal("Active"), StringVal("Score")},
				{StringVal("A"), BoolVal(true), NumberVal(10)},
				{StringVal("B"), BoolVal(false), NumberVal(20)},
				{StringVal("C"), BoolVal(true), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Active")},
				{BoolVal(false)},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  20, // only B
		},
		// --- large numbers ---
		{
			name: "min with large numbers",
			db: [][]Value{
				{StringVal("ID"), StringVal("Amount")},
				{StringVal("A"), NumberVal(1e10)},
				{StringVal("B"), NumberVal(1e12)},
				{StringVal("C"), NumberVal(1e8)},
			},
			crit:     [][]Value{{StringVal("ID")}, {StringVal("")}},
			field:    `"Amount"`,
			wantType: ValueNumber,
			wantNum:  1e8,
		},
		// --- zero values ---
		{
			name: "min when values include zero",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(10)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "min when zero is not the smallest",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(0)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(0)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  -5,
		},
		// --- single row database ---
		{
			name: "single row database returns that value",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Only"), NumberVal(99)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  99,
		},
		// --- Excel documentation example ---
		{
			name: "Excel doc example - min profit Apple h>10 age<16",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<16")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // Apple h=14, age=15 matches; h=18 age=20 excluded by <16
		},
		// --- mixed positive and negative ---
		{
			name: "min of mixed positive and negative",
			db: [][]Value{
				{StringVal("Type"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("A"), NumberVal(5)},
				{StringVal("A"), NumberVal(-3)},
				{StringVal("A"), NumberVal(2)},
			},
			crit:     [][]Value{{StringVal("Type")}, {StringVal("A")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -10,
		},
		// --- very close values (precision) ---
		{
			name: "min distinguishes very close values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.0000001)},
				{StringVal("B"), NumberVal(1.0000002)},
				{StringVal("C"), NumberVal(1.0000000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1.0000000,
		},
		// --- field as float column number ---
		{
			name: "field as float 5.5 truncates to column 5",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5.5",
			wantType: ValueNumber,
			wantNum:  45, // column 5 = Profit, min Apple = 45
		},
		// --- criteria on same column as field ---
		{
			name: "criteria on same column as field - min of Height where Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  12, // h=18,12,13,14 -> min=12
		},
		// --- = empty matches only empty cells ---
		{
			name: "= empty criteria matches empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(50)},
				{StringVal("C"), NumberVal(30)},
				{EmptyVal(), NumberVal(40)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  40, // min of 50 and 40
		},
		// --- <> empty matches non-empty cells ---
		{
			name: "<> empty criteria matches non-empty cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Score")},
				{StringVal("A"), NumberVal(10)},
				{EmptyVal(), NumberVal(5)},
				{StringVal("C"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("<>")},
			},
			field:    `"Score"`,
			wantType: ValueNumber,
			wantNum:  10, // min of 10 and 30 (empty Name row excluded)
		},
		// --- whitespace-padded criteria header ---
		{
			name: "whitespace-padded criteria header still matches",
			db:   db,
			crit: [][]Value{
				{StringVal(" Tree ")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		// --- empty database ---
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- duplicate criteria rows ---
		{
			name: "duplicate criteria rows still min correctly",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  76.8,
		},
		// --- multiple AND columns with OR rows ---
		{
			name: "multiple AND columns with multiple OR rows",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Yield")},
				{StringVal("Apple"), StringVal("<8")},
				{StringVal("Pear"), StringVal(">9")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45, // Apple y<8: Apple(y=6,p=45); Pear y>9: Pear(y=10,p=96) -> min=45
		},
		// --- field not found ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- field index errors ---
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index negative returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "-1",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- single positive value among text/empty ---
		{
			name: "single numeric value among text and empty",
			db: [][]Value{
				{StringVal("Cat"), StringVal("Val")},
				{StringVal("X"), StringVal("hello")},
				{StringVal("X"), NumberVal(42)},
				{StringVal("X"), EmptyVal()},
			},
			crit:     [][]Value{{StringVal("Cat")}, {StringVal("X")}},
			field:    `"Val"`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		// --- boolean column not considered for min ---
		{
			name: "boolean column returns 0 for min",
			db: [][]Value{
				{StringVal("Name"), StringVal("Flag")},
				{StringVal("A"), BoolVal(true)},
				{StringVal("B"), BoolVal(false)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Flag"`,
			wantType: ValueNumber,
			wantNum:  0, // booleans not numeric
		},
		// --- case-insensitive criteria ---
		{
			name: "case-insensitive text criteria match for min",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("APPLE")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		// --- min with only one positive and rest negative ---
		{
			name: "min where positives and negatives both present",
			db: [][]Value{
				{StringVal("Name"), StringVal("Val")},
				{StringVal("A"), NumberVal(100)},
				{StringVal("B"), NumberVal(-1)},
				{StringVal("C"), NumberVal(-200)},
				{StringVal("D"), NumberVal(50)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Val"`,
			wantType: ValueNumber,
			wantNum:  -200,
		},
	}

	runDBTests(t, "DMIN", tests)
}

func TestDMIN_WrongArgCount(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDMin(tt.args)
			if err != nil {
				t.Fatalf("fnDMin error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDMin(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDMIN_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the field column, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DMIN(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DMIN with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDMIN_ErrorInNonFieldColumn(t *testing.T) {
	// Error in a non-field column should not affect the result.
	db := [][]Value{
		{StringVal("Name"), StringVal("Score"), StringVal("Value")},
		{StringVal("A"), NumberVal(100), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValNA), NumberVal(20)},
		{StringVal("C"), NumberVal(300), NumberVal(30)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DMIN(A1:C4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("DMIN = %+v, want 10", got)
	}
}

// ---------------------------------------------------------------------------
// DPRODUCT
// ---------------------------------------------------------------------------

func TestDPRODUCT(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- basic: product of numeric column with single text criteria ---
		{
			name: "product Apple yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  840, // 14*10*6
		},
		{
			name: "product Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  80, // 10*8
		},
		{
			name: "product Cherry yield single record equals that value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  9,
		},

		// --- field by column number ---
		{
			name: "field by column number",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "4", // 4th column = Yield
			wantType: ValueNumber,
			wantNum:  840,
		},

		// --- no matching records returns 0 ---
		{
			name:     "no matches returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  0,
		},

		// --- blank criteria matches all records ---
		{
			name: "blank criteria matches all - product of all yields",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  604800, // 14*10*9*10*8*6
		},

		// --- no criteria rows matches all records ---
		{
			name: "no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  604800,
		},

		// --- multiple criteria rows = OR logic ---
		{
			name: "OR logic - Apple OR Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  67200, // 14*10*6 * 10*8 = 840*80
		},

		// --- multiple criteria columns = AND logic ---
		{
			name: "AND logic - Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  140, // Apple h=18 yield=14, Apple h=14 yield=10 → 14*10
		},

		// --- combined AND/OR criteria ---
		{
			name: "combined AND/OR - Apple Height>10 OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  11200, // (14*10) * (10*8) = 140*80
		},

		// --- numeric comparison criteria ---
		{
			name: "greater than on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">13")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  140, // h=18→14, h=14→10 → 14*10
		},
		{
			name: "less than on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  48, // h=9→8, h=8→6 → 8*6
		},
		{
			name: "greater or equal on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=13")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  1260, // h=18→14, h=13→9, h=14→10 → 14*9*10
		},
		{
			name: "less or equal on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  480, // h=12→10, h=9→8, h=8→6 → 10*8*6
		},
		{
			name: "not equal on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  67200, // all except Cherry (h=13): 14*10*10*8*6
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  10, // only Apple h=14, yield=10
		},

		// --- wildcard criteria ---
		{
			name: "wildcard * matches Apple",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  840, // 14*10*6
		},
		{
			name: "wildcard ? matches Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  80, // 10*8
		},

		// --- case-insensitive text criteria ---
		{
			name: "case-insensitive text match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  840,
		},

		// --- product with zero in data → 0 ---
		{
			name: "product with zero in data",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(5)},
				{StringVal("B"), NumberVal(0)},
				{StringVal("C"), NumberVal(3)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0, // 5*0*3
		},

		// --- product with negative numbers ---
		{
			name: "product with negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-3)},
				{StringVal("B"), NumberVal(4)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -12, // -3*4
		},
		{
			name: "product of two negatives is positive",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-2)},
				{StringVal("B"), NumberVal(-5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  10, // -2*-5
		},

		// --- product with all ones → 1 ---
		{
			name: "product all ones",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), NumberVal(1)},
				{StringVal("C"), NumberVal(1)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  1,
		},

		// --- decimal values ---
		{
			name: "product with decimal values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(2.5)},
				{StringVal("B"), NumberVal(4.0)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  10, // 2.5*4.0
		},

		// --- text column (non-numeric field values) → product ignores text ---
		{
			name: "text-only field column returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("foo")},
				{StringVal("B"), StringVal("bar")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},

		// --- mixed types: product ignores text, uses numbers ---
		{
			name: "mixed types - product ignores text cells",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(3)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  15, // 3*5
		},

		// --- empty cells in field column are ignored ---
		{
			name: "empty cells in field are ignored",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(7)},
				{StringVal("B"), EmptyVal()},
				{StringVal("C"), NumberVal(3)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  21, // 7*3
		},

		// --- large product values ---
		{
			name: "large product",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1000)},
				{StringVal("B"), NumberVal(2000)},
				{StringVal("C"), NumberVal(3000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  6e9, // 1000*2000*3000
		},

		// --- cross-column criteria ---
		{
			name: "cross-column criteria on different columns",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  7875, // Apple age=20 profit=105, Apple age=15 profit=75 → 105*75
		},

		// --- field name not found → #VALUE! ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},

		// --- field index out of range → #VALUE! ---
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},

		// --- field index 0 → #VALUE! ---
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},

		// --- product of profit column ---
		{
			name: "product of Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  354375, // 105*75*45
		},

		// --- Excel documentation example ---
		// DPRODUCT(database, "Yield", criteria) where criteria selects
		// Apple with Height>10 AND Height<16 (row 1) OR Pear (row 2).
		// Apple h=14, yield=10; Pear h=12 yield=10, h=9 yield=8 → 10*10*8=800
		{
			name: "Excel doc example - Apple h>10 h<16 OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10"), StringVal("<16")},
				{StringVal("Pear"), StringVal(""), StringVal("")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  800, // Apple h=14 yield=10; Pear yields 10,8 → 10*10*8
		},
	}

	runDBTests(t, "DPRODUCT", tests)
}

func TestDPRODUCT_WrongArgCount(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDProduct(tt.args)
			if err != nil {
				t.Fatalf("fnDProduct error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDProduct(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDPRODUCT_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the field column, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DPRODUCT(A1:B4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DPRODUCT with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDPRODUCT_ErrorInNonFieldColumn(t *testing.T) {
	// Error in a non-field column should not affect the result.
	db := [][]Value{
		{StringVal("Name"), StringVal("Score"), StringVal("Value")},
		{StringVal("A"), NumberVal(100), NumberVal(2)},
		{StringVal("B"), ErrorVal(ErrValNA), NumberVal(3)},
		{StringVal("C"), NumberVal(300), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := `DPRODUCT(A1:C4,"Value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("DPRODUCT = %+v, want 30", got)
	}
}

func TestDPRODUCT_FieldCaseInsensitive(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("VALUE")},
		{StringVal("A"), NumberVal(4)},
		{StringVal("B"), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	// Field name uses different case
	formula := `DPRODUCT(A1:B3,"value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("DPRODUCT case-insensitive field = %+v, want 20", got)
	}
}

// ---------------------------------------------------------------------------
// DPRODUCT — additional comprehensive tests
// ---------------------------------------------------------------------------

func TestDPRODUCT_SingleMatchReturnsValue(t *testing.T) {
	// Single matching record should return the value itself.
	db := [][]Value{
		{StringVal("Item"), StringVal("Qty")},
		{StringVal("A"), NumberVal(42)},
		{StringVal("B"), NumberVal(10)},
	}
	crit := [][]Value{
		{StringVal("Item")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,"Qty",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("DPRODUCT single match = %+v, want 42", got)
	}
}

func TestDPRODUCT_NoMatchReturnsZero(t *testing.T) {
	db := [][]Value{
		{StringVal("Item"), StringVal("Qty")},
		{StringVal("A"), NumberVal(5)},
		{StringVal("B"), NumberVal(3)},
	}
	crit := [][]Value{
		{StringVal("Item")},
		{StringVal("Z")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,"Qty",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("DPRODUCT no matches = %+v, want 0", got)
	}
}

func TestDPRODUCT_AllRecordsMatch(t *testing.T) {
	db := [][]Value{
		{StringVal("Cat"), StringVal("Val")},
		{StringVal("X"), NumberVal(2)},
		{StringVal("Y"), NumberVal(3)},
		{StringVal("Z"), NumberVal(7)},
	}
	// Empty criteria row means match all.
	crit := [][]Value{
		{StringVal("Cat")},
		{StringVal("")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 { // 2*3*7
		t.Errorf("DPRODUCT all match = %+v, want 42", got)
	}
}

func TestDPRODUCT_FieldByColumnNumber(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("A"), StringVal("B")},
		{StringVal("X"), NumberVal(3), NumberVal(10)},
		{StringVal("Y"), NumberVal(5), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	// Column 3 = "B"
	cf := evalCompile(t, `DPRODUCT(A1:C3,3,G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 { // 10*20
		t.Errorf("DPRODUCT field by col num = %+v, want 200", got)
	}
}

func TestDPRODUCT_NumericComparisonCriteria(t *testing.T) {
	db := [][]Value{
		{StringVal("Item"), StringVal("Price")},
		{StringVal("A"), NumberVal(3)},
		{StringVal("B"), NumberVal(7)},
		{StringVal("C"), NumberVal(12)},
		{StringVal("D"), NumberVal(4)},
	}

	tests := []struct {
		name string
		crit string
		want float64
	}{
		{"gt5", ">5", 84},      // 7*12 = 84
		{"lte4", "<=4", 12},    // 3*4 = 12
		{"gte7", ">=7", 84},    // 7*12 = 84
		{"lt7", "<7", 12},      // 3*4 = 12
		{"eq12", "=12", 12},    // 12
		{"ne7", "<>7", 144},    // 3*12*4 = 144
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			crit := [][]Value{
				{StringVal("Price")},
				{StringVal(tt.crit)},
			}
			resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
			cf := evalCompile(t, `DPRODUCT(A1:B5,"Price",G1:G2)`)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("DPRODUCT %s = %+v, want %v", tt.crit, got, tt.want)
			}
		})
	}
}

func TestDPRODUCT_MultipleCriteriaColumnsAND(t *testing.T) {
	db := [][]Value{
		{StringVal("Color"), StringVal("Size"), StringVal("Qty")},
		{StringVal("Red"), NumberVal(10), NumberVal(2)},
		{StringVal("Blue"), NumberVal(20), NumberVal(3)},
		{StringVal("Red"), NumberVal(30), NumberVal(5)},
		{StringVal("Red"), NumberVal(10), NumberVal(7)},
	}
	// AND: Color=Red AND Size=10
	crit := [][]Value{
		{StringVal("Color"), StringVal("Size")},
		{StringVal("Red"), NumberVal(10)},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:C5,"Qty",G1:H2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 14 { // 2*7
		t.Errorf("DPRODUCT AND criteria = %+v, want 14", got)
	}
}

func TestDPRODUCT_ORCriteriaMultipleRows(t *testing.T) {
	db := [][]Value{
		{StringVal("Fruit"), StringVal("Val")},
		{StringVal("Apple"), NumberVal(2)},
		{StringVal("Banana"), NumberVal(3)},
		{StringVal("Cherry"), NumberVal(5)},
	}
	// OR: Fruit=Apple OR Fruit=Cherry
	crit := [][]Value{
		{StringVal("Fruit")},
		{StringVal("Apple")},
		{StringVal("Cherry")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 { // 2*5
		t.Errorf("DPRODUCT OR criteria = %+v, want 10", got)
	}
}

func TestDPRODUCT_WildcardMatchStar(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Score")},
		{StringVal("Alpha"), NumberVal(2)},
		{StringVal("Beta"), NumberVal(3)},
		{StringVal("Alphabet"), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("Alph*")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Score",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 { // 2*5
		t.Errorf("DPRODUCT wildcard * = %+v, want 10", got)
	}
}

func TestDPRODUCT_WildcardMatchQuestion(t *testing.T) {
	db := [][]Value{
		{StringVal("Code"), StringVal("Val")},
		{StringVal("AB"), NumberVal(3)},
		{StringVal("AC"), NumberVal(4)},
		{StringVal("ABC"), NumberVal(100)},
	}
	crit := [][]Value{
		{StringVal("Code")},
		{StringVal("A?")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 12 { // 3*4
		t.Errorf("DPRODUCT wildcard ? = %+v, want 12", got)
	}
}

func TestDPRODUCT_CaseInsensitiveCriteria(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("Apple"), NumberVal(6)},
		{StringVal("APPLE"), NumberVal(7)},
		{StringVal("Banana"), NumberVal(100)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("apple")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 { // 6*7
		t.Errorf("DPRODUCT case-insensitive criteria = %+v, want 42", got)
	}
}

func TestDPRODUCT_NegativeValuesInProduct(t *testing.T) {
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(-2)},
		{StringVal("A"), NumberVal(3)},
		{StringVal("A"), NumberVal(-4)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 24 { // -2*3*-4 = 24
		t.Errorf("DPRODUCT neg values = %+v, want 24", got)
	}
}

func TestDPRODUCT_ZeroInProductYieldsZero(t *testing.T) {
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(100)},
		{StringVal("A"), NumberVal(0)},
		{StringVal("A"), NumberVal(50)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("DPRODUCT zero in product = %+v, want 0", got)
	}
}

func TestDPRODUCT_ErrorInMatchingCell(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(5)},
		{StringVal("A"), ErrorVal(ErrValNA)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("DPRODUCT error in cell = %+v, want #N/A", got)
	}
}

func TestDPRODUCT_VsDSUM_Relationship(t *testing.T) {
	// For a single matching record, DPRODUCT and DSUM should return the same value.
	db := [][]Value{
		{StringVal("Key"), StringVal("Val")},
		{StringVal("A"), NumberVal(17)},
		{StringVal("B"), NumberVal(99)},
	}
	crit := [][]Value{
		{StringVal("Key")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	cfProduct := evalCompile(t, `DPRODUCT(A1:B3,"Val",G1:G2)`)
	cfSum := evalCompile(t, `DSUM(A1:B3,"Val",G1:G2)`)

	gotProduct, err := Eval(cfProduct, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DPRODUCT: %v", err)
	}
	gotSum, err := Eval(cfSum, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSUM: %v", err)
	}

	if gotProduct.Type != ValueNumber || gotSum.Type != ValueNumber {
		t.Fatalf("expected both to be numbers, got product=%+v sum=%+v", gotProduct, gotSum)
	}
	if gotProduct.Num != gotSum.Num {
		t.Errorf("single match: DPRODUCT=%v DSUM=%v, expected equal", gotProduct.Num, gotSum.Num)
	}
}

func TestDPRODUCT_FractionalValues(t *testing.T) {
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(0.5)},
		{StringVal("A"), NumberVal(0.25)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0.125 { // 0.5*0.25
		t.Errorf("DPRODUCT fractional = %+v, want 0.125", got)
	}
}

func TestDPRODUCT_BoolFieldCoercedToColumnIndex(t *testing.T) {
	// TRUE coerces to 1, so field=TRUE means column 1.
	db := [][]Value{
		{StringVal("Val"), StringVal("Other")},
		{NumberVal(3), NumberVal(100)},
		{NumberVal(5), NumberVal(200)},
	}
	crit := [][]Value{
		{StringVal("Val")},
		{StringVal("")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,TRUE,G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 { // 3*5
		t.Errorf("DPRODUCT bool field = %+v, want 15", got)
	}
}

// ---------------------------------------------------------------------------
// DPRODUCT — extended comprehensive tests
// ---------------------------------------------------------------------------

func TestDPRODUCT_VsDGET_SingleMatch(t *testing.T) {
	// Cross-check: DPRODUCT with single match = DGET for same criteria.
	db := [][]Value{
		{StringVal("Key"), StringVal("Val")},
		{StringVal("A"), NumberVal(17)},
		{StringVal("B"), NumberVal(99)},
		{StringVal("C"), NumberVal(55)},
	}
	crit := [][]Value{
		{StringVal("Key")},
		{StringVal("B")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	cfProduct := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	cfGet := evalCompile(t, `DGET(A1:B4,"Val",G1:G2)`)

	gotProduct, err := Eval(cfProduct, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DPRODUCT: %v", err)
	}
	gotGet, err := Eval(cfGet, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DGET: %v", err)
	}

	if gotProduct.Type != ValueNumber || gotGet.Type != ValueNumber {
		t.Fatalf("expected both numbers, DPRODUCT=%+v DGET=%+v", gotProduct, gotGet)
	}
	if gotProduct.Num != gotGet.Num {
		t.Errorf("single match: DPRODUCT=%v DGET=%v, expected equal", gotProduct.Num, gotGet.Num)
	}
}

func TestDPRODUCT_ThreeMatchesProduct(t *testing.T) {
	// Three matching records: product = 2*5*7 = 70.
	db := [][]Value{
		{StringVal("Cat"), StringVal("Val")},
		{StringVal("A"), NumberVal(2)},
		{StringVal("B"), NumberVal(100)},
		{StringVal("A"), NumberVal(5)},
		{StringVal("A"), NumberVal(7)},
	}
	crit := [][]Value{
		{StringVal("Cat")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B5,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 70 {
		t.Errorf("DPRODUCT three matches = %+v, want 70", got)
	}
}

func TestDPRODUCT_VerySmallFractionalProduct(t *testing.T) {
	// Product of very small fractions: 0.1 * 0.01 * 0.001 = 1e-6.
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(0.1)},
		{StringVal("A"), NumberVal(0.01)},
		{StringVal("A"), NumberVal(0.001)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("DPRODUCT type = %v, want number", got.Type)
	}
	diff := got.Num - 1e-6
	if diff > 1e-15 || diff < -1e-15 {
		t.Errorf("DPRODUCT very small fractions = %g, want 1e-6", got.Num)
	}
}

func TestDPRODUCT_MixedNegativePositiveZero(t *testing.T) {
	// Mix: -3 * 4 * 0 = 0.
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(-3)},
		{StringVal("A"), NumberVal(4)},
		{StringVal("A"), NumberVal(0)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("DPRODUCT mixed neg/pos/zero = %+v, want 0", got)
	}
}

func TestDPRODUCT_CriteriaHeaderMismatch(t *testing.T) {
	// Criteria header does not match any database column.
	// Non-blank criterion on unmatched header -> no match.
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("NonExistentCol")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B3,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("DPRODUCT unmatched criteria header = %+v, want 0", got)
	}
}

func TestDPRODUCT_EmptyFieldArgError(t *testing.T) {
	// Empty field argument -> #VALUE! error.
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(10)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}
	args := []Value{
		{Type: ValueArray, Array: db},
		EmptyVal(),
		{Type: ValueArray, Array: crit},
	}
	got, err := fnDProduct(args)
	if err != nil {
		t.Fatalf("fnDProduct error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("DPRODUCT empty field = %+v, want #VALUE!", got)
	}
}

func TestDPRODUCT_ErrorInDatabaseArg(t *testing.T) {
	// Error value as database argument -> propagate error.
	args := []Value{
		ErrorVal(ErrValREF),
		StringVal("Val"),
		{Type: ValueArray, Array: [][]Value{{StringVal("X")}, {StringVal("A")}}},
	}
	got, err := fnDProduct(args)
	if err != nil {
		t.Fatalf("fnDProduct error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("DPRODUCT error db arg = %+v, want #REF!", got)
	}
}

func TestDPRODUCT_ErrorInCriteriaArg(t *testing.T) {
	// Error value as criteria argument -> propagate error.
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(10)},
	}
	args := []Value{
		{Type: ValueArray, Array: db},
		StringVal("Val"),
		ErrorVal(ErrValNA),
	}
	got, err := fnDProduct(args)
	if err != nil {
		t.Fatalf("fnDProduct error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("DPRODUCT error criteria arg = %+v, want #N/A", got)
	}
}

func TestDPRODUCT_MultipleORWithAND(t *testing.T) {
	// Two OR rows each with AND conditions:
	// Row 1: Color=Red AND Size>15
	// Row 2: Color=Blue AND Size<=10
	db := [][]Value{
		{StringVal("Color"), StringVal("Size"), StringVal("Qty")},
		{StringVal("Red"), NumberVal(20), NumberVal(2)},   // matches row 1
		{StringVal("Red"), NumberVal(10), NumberVal(3)},   // no match
		{StringVal("Blue"), NumberVal(5), NumberVal(4)},   // matches row 2
		{StringVal("Blue"), NumberVal(30), NumberVal(5)},  // no match
		{StringVal("Green"), NumberVal(8), NumberVal(6)},  // no match
	}
	crit := [][]Value{
		{StringVal("Color"), StringVal("Size")},
		{StringVal("Red"), StringVal(">15")},
		{StringVal("Blue"), StringVal("<=10")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:C6,"Qty",G1:H3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 8 { // 2*4
		t.Errorf("DPRODUCT multiple OR with AND = %+v, want 8", got)
	}
}

func TestDPRODUCT_SingleRecordDB(t *testing.T) {
	// Database with only one record (plus header).
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("Solo"), NumberVal(42)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("Solo")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B2,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("DPRODUCT single record DB = %+v, want 42", got)
	}
}

func TestDPRODUCT_LargeNegativeProduct(t *testing.T) {
	// Odd number of negative values: -2 * -3 * -5 = -30.
	db := [][]Value{
		{StringVal("X"), StringVal("Y")},
		{StringVal("A"), NumberVal(-2)},
		{StringVal("A"), NumberVal(-3)},
		{StringVal("A"), NumberVal(-5)},
	}
	crit := [][]Value{
		{StringVal("X")},
		{StringVal("A")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Y",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -30 {
		t.Errorf("DPRODUCT odd negatives = %+v, want -30", got)
	}
}

func TestDPRODUCT_NotEqualStringCriteria(t *testing.T) {
	// "<>Apple" should match everything except Apple.
	db := [][]Value{
		{StringVal("Fruit"), StringVal("Val")},
		{StringVal("Apple"), NumberVal(2)},
		{StringVal("Banana"), NumberVal(3)},
		{StringVal("Cherry"), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Fruit")},
		{StringVal("<>Apple")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 { // 3*5
		t.Errorf("DPRODUCT <>Apple = %+v, want 15", got)
	}
}

func TestDPRODUCT_WildcardMiddleStar(t *testing.T) {
	// Wildcard with star in the middle: "A*e" matches "Apple", "Ape".
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("Apple"), NumberVal(2)},
		{StringVal("Ape"), NumberVal(3)},
		{StringVal("Banana"), NumberVal(100)},
		{StringVal("Axe"), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("A*e")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B5,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 { // 2*3*5
		t.Errorf("DPRODUCT wildcard middle star = %+v, want 30", got)
	}
}

func TestDPRODUCT_BoolCriteriaValue(t *testing.T) {
	// Criteria matching boolean values in the database.
	db := [][]Value{
		{StringVal("Active"), StringVal("Val")},
		{BoolVal(true), NumberVal(3)},
		{BoolVal(false), NumberVal(7)},
		{BoolVal(true), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Active")},
		{BoolVal(true)},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 { // 3*5
		t.Errorf("DPRODUCT bool criteria = %+v, want 15", got)
	}
}

func TestDPRODUCT_FieldErrorArg(t *testing.T) {
	// Error value as field argument -> propagate error.
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(10)},
	}
	args := []Value{
		{Type: ValueArray, Array: db},
		ErrorVal(ErrValNUM),
		{Type: ValueArray, Array: [][]Value{{StringVal("Name")}, {StringVal("A")}}},
	}
	got, err := fnDProduct(args)
	if err != nil {
		t.Fatalf("fnDProduct error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("DPRODUCT error field arg = %+v, want #NUM!", got)
	}
}

func TestDPRODUCT_NoCriteriaRowsMatchAll(t *testing.T) {
	// Criteria with only header row and no condition rows -> match all.
	db := [][]Value{
		{StringVal("Name"), StringVal("Val")},
		{StringVal("A"), NumberVal(2)},
		{StringVal("B"), NumberVal(3)},
		{StringVal("C"), NumberVal(5)},
	}
	crit := [][]Value{
		{StringVal("Name")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	cf := evalCompile(t, `DPRODUCT(A1:B4,"Val",G1:G1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 { // 2*3*5
		t.Errorf("DPRODUCT no criteria rows = %+v, want 30", got)
	}
}

// ---------------------------------------------------------------------------
// DSTDEV
// ---------------------------------------------------------------------------

func TestDSTDEV(t *testing.T) {
	db := standardDB()

	// Apple profits: 105, 75, 45. mean=75, ss=1800, sample_var=900, stdev=30
	// Pear profits: 96, 76.8. mean=86.4, ss=184.32, sample_var=184.32, stdev=sqrt(184.32)
	// All profits: 105, 96, 105, 75, 76.8, 45. mean=83.8
	//   ss = 449.44+148.84+449.44+77.44+49+1504.84 = 2679, var=535.8, stdev=sqrt(535.8)

	tests := []dbTestCase{
		// --- Basic: stdev of numeric column with single text criteria ---
		{
			name: "basic stdev Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  30, // sample stdev
		},
		{
			name: "basic stdev Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Pear profits: 96, 76.8. mean=86.4, var=((9.6^2+9.6^2)/1)=184.32
			wantNum: math.Sqrt(184.32),
		},
		// --- Field specified by column number ---
		{
			name: "field by column number",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // 5th column = Profit
			wantType: ValueNumber,
			wantNum:  30,
		},
		{
			name: "field by column number 4 yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "4", // 4th column = Yield
			wantType: ValueNumber,
			// Apple yields: 14, 10, 6. mean=10, ss=32, var=16, stdev=4
			wantNum: 4,
		},
		// --- Multiple criteria rows (OR logic) ---
		{
			name: "OR criteria Apple or Pear yield matches Excel docs",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			// Apple+Pear yields: 14, 10, 6, 10, 8. mean=9.6
			// ss=35.2, var=8.8, stdev=sqrt(8.8)≈2.96648
			wantNum: math.Sqrt(8.8),
		},
		{
			name: "OR criteria Apple or Cherry profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// 105, 105, 75, 45. mean=82.5, ss=2475, var=825, stdev=sqrt(825)
			wantNum: math.Sqrt(825),
		},
		// --- Multiple criteria columns (AND logic) ---
		{
			name: "AND criteria Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105 (h=18), 75 (h=14). mean=90, ss=450, var=450, stdev=sqrt(450)
			wantNum: math.Sqrt(450),
		},
		{
			name: "AND criteria Apple with Age>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple age>10: 105 (age=20), 75 (age=15). mean=90, stdev=sqrt(450)
			wantNum: math.Sqrt(450),
		},
		// --- No matching records → #DIV/0! ---
		{
			name:     "no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- All records match (empty/blank criteria) ---
		{
			name: "blank criteria matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")}, // blank = match all
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// All profits: 105, 96, 105, 75, 76.8, 45. mean=502.8/6=83.8
			// ss=449.44+148.84+449.44+77.44+49+1505.44=2679.6
			// var=2679.6/5=535.92, stdev=sqrt(535.92)
			wantNum: math.Sqrt(2679.6 / 5),
		},
		// --- Single record match → #DIV/0! ---
		{
			name:     "single value returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Two records match → computable ---
		{
			name: "two records are sufficient",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			// Pear heights: 12, 9. mean=10.5, ss=4.5, var=4.5, stdev=sqrt(4.5)
			wantNum: math.Sqrt(4.5),
		},
		// --- Numeric criteria: > ---
		{
			name: "numeric criteria greater than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>12: 18→105, 13→105, 14→75. mean=95, ss=600, var=300, stdev=sqrt(300)
			wantNum: math.Sqrt(300),
		},
		// --- Numeric criteria: < ---
		{
			name: "numeric criteria less than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<12: 9→76.8, 8→45. mean=60.9, ss=505.62, var=505.62, stdev=sqrt(505.62)
			wantNum: math.Sqrt(505.62),
		},
		// --- Numeric criteria: >= ---
		{
			name: "numeric criteria greater than or equal",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>=13: 18→105, 13→105, 14→75. mean=95, ss=600, var=300, stdev=sqrt(300)
			wantNum: math.Sqrt(300),
		},
		// --- Numeric criteria: <= ---
		{
			name: "numeric criteria less than or equal",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<=12: 12→96, 9→76.8, 8→45. mean=72.6
			// ss = (23.4^2 + 4.2^2 + 27.6^2) = 547.56+17.64+761.76 = 1326.96
			// var=663.48, stdev=sqrt(663.48)
			wantNum: math.Sqrt(1326.96 / 2),
		},
		// --- Numeric criteria: <> ---
		{
			name: "numeric criteria not equal",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<>13: 105, 96, 75, 76.8, 45. mean=397.8/5=79.56
			// ss = (25.44^2+16.44^2+4.56^2+2.76^2+34.56^2)
			// = 647.1936+270.2736+20.7936+7.6176+1194.3936 = 2140.272
			// var=2140.272/4=535.068, stdev=sqrt(535.068)
			wantNum: math.Sqrt(2140.272 / 4),
		},
		// --- Wildcard criteria: * ---
		{
			name: "wildcard star criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")}, // matches Apple
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  30, // same as Apple stdev
		},
		// --- Wildcard criteria: ? ---
		{
			name: "wildcard question mark criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")}, // matches Pear
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(184.32),
		},
		// --- Cross-column criteria (complex OR with AND) ---
		{
			name: "cross-column OR with AND: Apple h>10 OR Pear h<10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. Pear h<10: 76.8. values: 105, 75, 76.8
			// mean = 256.8/3 = 85.6
			// ss = (19.4^2+10.6^2+8.8^2) = 376.36+112.36+77.44 = 566.16
			// var = 566.16/2 = 283.08, stdev=sqrt(283.08)
			wantNum: math.Sqrt(283.08),
		},
		// --- All same values → 0 ---
		{
			name: "all same values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(5)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "all same values two records returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Negative numbers ---
		{
			name: "negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(-5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// values: -10, -20, -5. mean=-35/3≈-11.667
			// ss = (1.667^2 + 8.333^2 + 6.667^2) = 2.7789+69.4389+44.4489 = 116.6667
			// var = 116.6667/2 = 58.3333, stdev=sqrt(58.3333)
			wantNum: math.Sqrt(350.0 / 6.0),
		},
		// --- Decimal values ---
		{
			name: "decimal values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.5)},
				{StringVal("C"), NumberVal(3.5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// values: 1.5, 2.5, 3.5. mean=2.5
			// ss = 1+0+1 = 2, var=1, stdev=1
			wantNum: 1,
		},
		// --- Text column → #DIV/0! (no numeric values to compute stdev) ---
		{
			name: "text column returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Label")},
				{StringVal("A"), StringVal("foo")},
				{StringVal("B"), StringVal("bar")},
				{StringVal("C"), StringVal("baz")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Label"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Relationship: DSTDEV > DSTDEVP for same data ---
		// (tested programmatically below in TestDSTDEV_GreaterThanDSTDEVP)

		// --- Field name error cases ---
		{
			name:     "field name not found returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index 0 returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- Known Excel result from documentation ---
		{
			name: "Excel docs example: stdev yield for Apple or Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			// Excel returns 2.96648 (=sqrt(8.8))
			wantNum: math.Sqrt(8.8),
		},
		// --- Error propagation ---
		{
			name: "error in field column propagates",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), ErrorVal(ErrValDIV0)},
				{StringVal("C"), NumberVal(20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Mixed types: only numbers used ---
		{
			name: "mixed types only uses numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), NumberVal(30)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// numeric values: 10, 20, 30. mean=20, ss=200, var=100, stdev=10
			wantNum: 10,
		},
		// --- Empty database → #DIV/0! ---
		{
			name: "empty database returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				// no data rows
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Large spread ---
		{
			name: "large spread values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), NumberVal(1000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// values: 1, 1000. mean=500.5, ss=2*499.5^2=499000.5
			// var=499000.5/1=499000.5, stdev=sqrt(499000.5)
			wantNum: math.Sqrt(499000.5),
		},
	}

	runDBTests(t, "DSTDEV", tests)
}

// TestDSTDEV_GreaterThanDSTDEVP verifies that DSTDEV > DSTDEVP for the same data,
// since sample standard deviation (n-1 denominator) is always larger than
// population standard deviation (n denominator) when n > 1.
func TestDSTDEV_GreaterThanDSTDEVP(t *testing.T) {
	db := standardDB()
	crit := [][]Value{
		{StringVal("Tree")},
		{StringVal("Apple")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	stdevFormula := `DSTDEV(A1:E7,"Profit",G1:G2)`
	stdevpFormula := `DSTDEVP(A1:E7,"Profit",G1:G2)`

	cfS := evalCompile(t, stdevFormula)
	gotS, err := Eval(cfS, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSTDEV: %v", err)
	}

	cfP := evalCompile(t, stdevpFormula)
	gotP, err := Eval(cfP, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSTDEVP: %v", err)
	}

	if gotS.Type != ValueNumber || gotP.Type != ValueNumber {
		t.Fatalf("expected both numeric, got DSTDEV=%+v DSTDEVP=%+v", gotS, gotP)
	}
	if gotS.Num <= gotP.Num {
		t.Errorf("DSTDEV (%g) should be > DSTDEVP (%g) for same data", gotS.Num, gotP.Num)
	}
}

func TestDSTDEV_WrongArgCount(t *testing.T) {
	result, err := fnDStdev(nil)
	if err != nil {
		t.Fatalf("fnDStdev error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDStdev(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DSTDEVP
// ---------------------------------------------------------------------------

func TestDSTDEVP(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- Basic tests ---
		{
			name: "stdevp Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple profits: 105,75,45. mean=75, ss=1800, var=600, stdevp=sqrt(600)
			wantNum: math.Sqrt(600),
		},
		{
			name: "stdevp single value returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name:     "stdevp no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "stdevp Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Pear: 96,76.8. mean=86.4, ss=184.32, var=92.16, stdevp=sqrt(92.16)
			wantNum: math.Sqrt(92.16),
		},
		{
			name: "stdevp equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(7)},
				{StringVal("B"), NumberVal(7)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Field by column number ---
		{
			name: "stdevp field by column number",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // Profit
			wantType: ValueNumber,
			wantNum:  math.Sqrt(600),
		},
		{
			name: "stdevp field by column number 4 yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "4", // Yield
			wantType: ValueNumber,
			// Apple yields: 14, 10, 6. mean=10, ss=32, var=32/3, stdevp=sqrt(32/3)
			wantNum: math.Sqrt(32.0 / 3.0),
		},
		// --- All records match (blank criteria) ---
		{
			name: "stdevp blank criteria matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// All profits: 105, 96, 105, 75, 76.8, 45. mean=502.8/6=83.8
			// ss=2679.6, var=2679.6/6=446.6, stdevp=sqrt(446.6)
			wantNum: math.Sqrt(2679.6 / 6),
		},
		{
			name: "stdevp no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(2679.6 / 6),
		},
		// --- Multiple criteria rows (OR logic) ---
		{
			name: "stdevp OR criteria Apple or Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			// Apple+Pear yields: 14, 10, 6, 10, 8. mean=9.6
			// ss=35.2, var=35.2/5=7.04, stdevp=sqrt(7.04)
			wantNum: math.Sqrt(7.04),
		},
		{
			name: "stdevp OR criteria Apple or Cherry profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// 105, 105, 75, 45. mean=82.5, ss=2475, var=2475/4=618.75
			wantNum: math.Sqrt(618.75),
		},
		// --- Multiple criteria columns (AND logic) ---
		{
			name: "stdevp AND criteria Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105 (h=18), 75 (h=14). mean=90, ss=450, var=450/2=225
			wantNum: math.Sqrt(225),
		},
		// --- Numeric criteria operators ---
		{
			name: "stdevp numeric criteria greater than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>12: 18→105, 13→105, 14→75. mean=95, ss=600, var=200
			wantNum: math.Sqrt(200),
		},
		{
			name: "stdevp numeric criteria less than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<12: 9→76.8, 8→45. mean=60.9, ss=505.62, var=252.81
			wantNum: math.Sqrt(252.81),
		},
		{
			name: "stdevp numeric criteria >=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>=13: 105, 105, 75. mean=95, ss=600, var=200
			wantNum: math.Sqrt(200),
		},
		{
			name: "stdevp numeric criteria <=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<=12: 96, 76.8, 45. mean=72.6, ss=1326.96, var=442.32
			wantNum: math.Sqrt(1326.96 / 3),
		},
		{
			name: "stdevp numeric criteria <>",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<>13: 105, 96, 75, 76.8, 45. mean=79.56, ss=2140.272, var=428.0544
			wantNum: math.Sqrt(2140.272 / 5),
		},
		// --- Wildcard criteria ---
		{
			name: "stdevp wildcard star criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(600), // same as Apple
		},
		{
			name: "stdevp wildcard question mark criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(92.16),
		},
		// --- Case insensitive matching ---
		{
			name: "stdevp case insensitive criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(600),
		},
		{
			name: "stdevp case insensitive field name",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  math.Sqrt(600),
		},
		// --- Error propagation ---
		{
			name: "stdevp error in field column propagates",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), ErrorVal(ErrValDIV0)},
				{StringVal("C"), NumberVal(20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Mixed types: only numbers used ---
		{
			name: "stdevp mixed types only uses numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), NumberVal(30)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// 10, 20, 30. mean=20, ss=200, var=200/3
			wantNum: math.Sqrt(200.0 / 3.0),
		},
		// --- Text column returns DIV/0 ---
		{
			name: "stdevp text column returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Label")},
				{StringVal("A"), StringVal("foo")},
				{StringVal("B"), StringVal("bar")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Label"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- All same values returns 0 ---
		{
			name: "stdevp all same values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(5)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Negative numbers ---
		{
			name: "stdevp negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(-5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=-35/3, ss=350/3, var=350/9, stdevp=sqrt(350/9)
			wantNum: math.Sqrt(350.0 / 9.0),
		},
		// --- Decimal values ---
		{
			name: "stdevp decimal values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.5)},
				{StringVal("C"), NumberVal(3.5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=2.5, ss=2, var=2/3, stdevp=sqrt(2/3)
			wantNum: math.Sqrt(2.0 / 3.0),
		},
		// --- Empty database ---
		{
			name: "stdevp empty database returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Large spread ---
		{
			name: "stdevp large spread values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), NumberVal(1000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=500.5, ss=499000.5, var=499000.5/2=249500.25
			wantNum: math.Sqrt(249500.25),
		},
		// --- Cross-column OR with AND ---
		{
			name: "stdevp cross-column OR with AND",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. Pear h<10: 76.8. values: 105, 75, 76.8
			// mean=85.6, ss=566.16, var=566.16/3=188.72
			wantNum: math.Sqrt(566.16 / 3),
		},
		// --- Field name not found ---
		{
			name:     "stdevp field name not found",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- Field index out of range ---
		{
			name:     "stdevp field index out of range",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "stdevp field index 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
	}

	runDBTests(t, "DSTDEVP", tests)
}

func TestDSTDEVP_WrongArgCount(t *testing.T) {
	result, err := fnDStdevP(nil)
	if err != nil {
		t.Fatalf("fnDStdevP error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDStdevP(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DVAR
// ---------------------------------------------------------------------------

func TestDVAR(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- Basic tests ---
		{
			name: "var Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  900, // sample var: (30^2+0+30^2)/2 = 900
		},
		{
			name: "var Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  184.32, // (9.6^2+9.6^2)/1
		},
		// --- Single value returns DIV/0 (sample needs n>=2) ---
		{
			name:     "var single value returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- No matching records ---
		{
			name:     "var no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Equal values ---
		{
			name: "var equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(10)},
				{StringVal("C"), NumberVal(10)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Field by column number ---
		{
			name: "var field by column number",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // Profit
			wantType: ValueNumber,
			wantNum:  900,
		},
		{
			name: "var field by column number 4 yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "4", // Yield
			wantType: ValueNumber,
			// Apple yields: 14, 10, 6. mean=10, ss=32, var=32/2=16
			wantNum: 16,
		},
		// --- All records match (blank criteria) ---
		{
			name: "var blank criteria matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// All profits: 105, 96, 105, 75, 76.8, 45. mean=83.8
			// ss=2679.6, var=2679.6/5=535.92
			wantNum: 2679.6 / 5,
		},
		{
			name: "var no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2679.6 / 5,
		},
		// --- Multiple criteria rows (OR logic) ---
		{
			name: "var OR criteria Apple or Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			// Apple+Pear yields: 14, 10, 6, 10, 8. mean=9.6
			// ss=35.2, var=35.2/4=8.8
			wantNum: 8.8,
		},
		{
			name: "var OR criteria Apple or Cherry profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// 105, 105, 75, 45. mean=82.5, ss=2475, var=2475/3=825
			wantNum: 825,
		},
		// --- Multiple criteria columns (AND logic) ---
		{
			name: "var AND criteria Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. mean=90, ss=450, var=450/1=450
			wantNum: 450,
		},
		{
			name: "var AND criteria Apple with Age>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple age>10: 105 (age=20), 75 (age=15). var=450
			wantNum: 450,
		},
		// --- Numeric criteria operators ---
		{
			name: "var numeric criteria greater than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>12: 105, 105, 75. mean=95, ss=600, var=300
			wantNum: 300,
		},
		{
			name: "var numeric criteria less than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<12: 76.8, 45. mean=60.9, ss=505.62, var=505.62
			wantNum: 505.62,
		},
		{
			name: "var numeric criteria >=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>=13: 105, 105, 75. mean=95, ss=600, var=300
			wantNum: 300,
		},
		{
			name: "var numeric criteria <=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<=12: 96, 76.8, 45. mean=72.6, ss=1326.96, var=663.48
			wantNum: 1326.96 / 2,
		},
		{
			name: "var numeric criteria <>",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<>13: 105, 96, 75, 76.8, 45. mean=79.56, ss=2140.272, var=535.068
			wantNum: 2140.272 / 4,
		},
		// --- Wildcard criteria ---
		{
			name: "var wildcard star criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  900, // same as Apple
		},
		{
			name: "var wildcard question mark criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  184.32,
		},
		// --- Case insensitive matching ---
		{
			name: "var case insensitive criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  900,
		},
		{
			name: "var case insensitive field name",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  900,
		},
		// --- Error propagation ---
		{
			name: "var error in field column propagates",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), ErrorVal(ErrValNUM)},
				{StringVal("C"), NumberVal(20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		// --- Mixed types: only numbers used ---
		{
			name: "var mixed types only uses numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), NumberVal(30)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// 10, 20, 30. mean=20, ss=200, var=100
			wantNum: 100,
		},
		// --- Text column returns DIV/0 ---
		{
			name: "var text column returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Label")},
				{StringVal("A"), StringVal("foo")},
				{StringVal("B"), StringVal("bar")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Label"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Negative numbers ---
		{
			name: "var negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(-5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=-35/3, ss=350/3, var=350/6
			wantNum: 350.0 / 6.0,
		},
		// --- Decimal values ---
		{
			name: "var decimal values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.5)},
				{StringVal("C"), NumberVal(3.5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=2.5, ss=2, var=1
			wantNum: 1,
		},
		// --- Empty database ---
		{
			name: "var empty database returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Large spread ---
		{
			name: "var large spread values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), NumberVal(1000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=500.5, ss=499000.5, var=499000.5
			wantNum: 499000.5,
		},
		// --- Cross-column OR with AND ---
		{
			name: "var cross-column OR with AND",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. Pear h<10: 76.8. values: 105, 75, 76.8
			// mean=85.6, ss=566.16, var=566.16/2=283.08
			wantNum: 283.08,
		},
		// --- Field errors ---
		{
			name:     "var field name not found",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "var field index out of range",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "var field index 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- Two equal values ---
		{
			name: "var two equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Known cross-check with Excel ---
		{
			name: "var all profits sample variance",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			// Heights: 18, 12, 13, 14, 9, 8. mean=74/6=12.333...
			// ss = 32.111+0.111+0.444+2.778+11.111+18.778 = 65.333
			// var = 65.333/5 = 13.0667
			wantNum: func() float64 {
				vals := []float64{18, 12, 13, 14, 9, 8}
				mean := 0.0
				for _, v := range vals {
					mean += v
				}
				mean /= float64(len(vals))
				ss := 0.0
				for _, v := range vals {
					d := v - mean
					ss += d * d
				}
				return ss / float64(len(vals)-1)
			}(),
		},
	}

	runDBTests(t, "DVAR", tests)
}

func TestDVAR_WrongArgCount(t *testing.T) {
	result, err := fnDVar(nil)
	if err != nil {
		t.Fatalf("fnDVar error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDVar(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DVARP
// ---------------------------------------------------------------------------

func TestDVARP(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		// --- Basic tests ---
		{
			name: "varp Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  600, // population var: (30^2+0+30^2)/3 = 600
		},
		{
			name: "varp Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  92.16, // (9.6^2+9.6^2)/2
		},
		// --- Single value returns 0 (population with n=1 is 0) ---
		{
			name: "varp single value returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- No matching records ---
		{
			name:     "varp no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Equal values ---
		{
			name: "varp equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(3)},
				{StringVal("B"), NumberVal(3)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Field by column number ---
		{
			name: "varp field by column number",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // Profit
			wantType: ValueNumber,
			wantNum:  600,
		},
		{
			name: "varp field by column number 4 yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "4", // Yield
			wantType: ValueNumber,
			// Apple yields: 14, 10, 6. mean=10, ss=32, var=32/3
			wantNum: 32.0 / 3.0,
		},
		// --- All records match (blank criteria) ---
		{
			name: "varp blank criteria matches all records",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// All profits: 105, 96, 105, 75, 76.8, 45. ss=2679.6, var=2679.6/6=446.6
			wantNum: 2679.6 / 6,
		},
		{
			name: "varp no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  2679.6 / 6,
		},
		// --- Multiple criteria rows (OR logic) ---
		{
			name: "varp OR criteria Apple or Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			// Apple+Pear yields: 14, 10, 6, 10, 8. mean=9.6
			// ss=35.2, var=35.2/5=7.04
			wantNum: 7.04,
		},
		{
			name: "varp OR criteria Apple or Cherry profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// 105, 105, 75, 45. mean=82.5, ss=2475, var=2475/4=618.75
			wantNum: 618.75,
		},
		// --- Multiple criteria columns (AND logic) ---
		{
			name: "varp AND criteria Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. mean=90, ss=450, var=450/2=225
			wantNum: 225,
		},
		// --- Numeric criteria operators ---
		{
			name: "varp numeric criteria greater than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>12: 105, 105, 75. mean=95, ss=600, var=200
			wantNum: 200,
		},
		{
			name: "varp numeric criteria less than",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<12: 76.8, 45. mean=60.9, ss=505.62, var=252.81
			wantNum: 252.81,
		},
		{
			name: "varp numeric criteria >=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h>=13: 105, 105, 75. mean=95, ss=600, var=200
			wantNum: 200,
		},
		{
			name: "varp numeric criteria <=",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<=12: 96, 76.8, 45. mean=72.6, ss=1326.96, var=442.32
			wantNum: 1326.96 / 3,
		},
		{
			name: "varp numeric criteria <>",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// h<>13: 105, 96, 75, 76.8, 45. mean=79.56, ss=2140.272, var=428.0544
			wantNum: 2140.272 / 5,
		},
		// --- Wildcard criteria ---
		{
			name: "varp wildcard star criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  600, // same as Apple
		},
		{
			name: "varp wildcard question mark criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  92.16,
		},
		// --- Case insensitive matching ---
		{
			name: "varp case insensitive criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  600,
		},
		{
			name: "varp case insensitive field name",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"profit"`,
			wantType: ValueNumber,
			wantNum:  600,
		},
		// --- Error propagation ---
		{
			name: "varp error in field column propagates",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), ErrorVal(ErrValNAME)},
				{StringVal("C"), NumberVal(20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValNAME,
		},
		// --- Mixed types: only numbers used ---
		{
			name: "varp mixed types only uses numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), NumberVal(30)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// 10, 20, 30. mean=20, ss=200, var=200/3
			wantNum: 200.0 / 3.0,
		},
		// --- Text column returns DIV/0 ---
		{
			name: "varp text column returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Label")},
				{StringVal("A"), StringVal("foo")},
				{StringVal("B"), StringVal("bar")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Label"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Negative numbers ---
		{
			name: "varp negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-20)},
				{StringVal("C"), NumberVal(-5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=-35/3, ss=350/3, var=350/9
			wantNum: 350.0 / 9.0,
		},
		// --- Decimal values ---
		{
			name: "varp decimal values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1.5)},
				{StringVal("B"), NumberVal(2.5)},
				{StringVal("C"), NumberVal(3.5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=2.5, ss=2, var=2/3
			wantNum: 2.0 / 3.0,
		},
		// --- Empty database ---
		{
			name: "varp empty database returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		// --- Large spread ---
		{
			name: "varp large spread values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(1)},
				{StringVal("B"), NumberVal(1000)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			// mean=500.5, ss=499000.5, var=249500.25
			wantNum: 249500.25,
		},
		// --- Cross-column OR with AND ---
		{
			name: "varp cross-column OR with AND",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple h>10: 105, 75. Pear h<10: 76.8. mean=85.6, ss=566.16, var=188.72
			wantNum: 566.16 / 3,
		},
		// --- Field errors ---
		{
			name:     "varp field name not found",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "varp field index out of range",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "varp field index 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		// --- Three equal values ---
		{
			name: "varp three equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(42)},
				{StringVal("B"), NumberVal(42)},
				{StringVal("C"), NumberVal(42)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		// --- Known cross-check: all heights ---
		{
			name: "varp all heights population variance",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum: func() float64 {
				vals := []float64{18, 12, 13, 14, 9, 8}
				mean := 0.0
				for _, v := range vals {
					mean += v
				}
				mean /= float64(len(vals))
				ss := 0.0
				for _, v := range vals {
					d := v - mean
					ss += d * d
				}
				return ss / float64(len(vals))
			}(),
		},
	}

	runDBTests(t, "DVARP", tests)
}

func TestDVARP_WrongArgCount(t *testing.T) {
	result, err := fnDVarP(nil)
	if err != nil {
		t.Fatalf("fnDVarP error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDVarP(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// Cross-verification tests: relationships between DVAR/DVARP/DSTDEV/DSTDEVP
// ---------------------------------------------------------------------------

// TestDVAR_Equals_DSTDEV_Squared verifies that DVAR = DSTDEV^2 for the same data.
func TestDVAR_Equals_DSTDEV_Squared(t *testing.T) {
	scenarios := []struct {
		name string
		crit [][]Value
	}{
		{
			name: "Apple profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
		},
		{
			name: "Pear profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Pear")}},
		},
		{
			name: "all records",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
		},
		{
			name: "Height > 10",
			crit: [][]Value{{StringVal("Height")}, {StringVal(">10")}},
		},
		{
			name: "Apple or Cherry",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}, {StringVal("Cherry")}},
		},
	}

	db := standardDB()
	for _, sc := range scenarios {
		t.Run(sc.name, func(t *testing.T) {
			resolver := makeDBResolver(db, 1, 1, sc.crit, 7, 1)
			critRange := dbRange(7, 1, len(sc.crit[0]), len(sc.crit))

			varFormula := `DVAR(A1:E7,"Profit",` + critRange + `)`
			stdevFormula := `DSTDEV(A1:E7,"Profit",` + critRange + `)`

			cfV := evalCompile(t, varFormula)
			gotV, err := Eval(cfV, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DVAR: %v", err)
			}

			cfS := evalCompile(t, stdevFormula)
			gotS, err := Eval(cfS, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DSTDEV: %v", err)
			}

			if gotV.Type != ValueNumber || gotS.Type != ValueNumber {
				// Both might be errors (e.g., single value). Skip.
				return
			}

			stdevSquared := gotS.Num * gotS.Num
			diff := gotV.Num - stdevSquared
			if diff > 1e-6 || diff < -1e-6 {
				t.Errorf("DVAR (%g) != DSTDEV^2 (%g), diff=%g", gotV.Num, stdevSquared, diff)
			}
		})
	}
}

// TestDVARP_Equals_DSTDEVP_Squared verifies that DVARP = DSTDEVP^2 for the same data.
func TestDVARP_Equals_DSTDEVP_Squared(t *testing.T) {
	scenarios := []struct {
		name string
		crit [][]Value
	}{
		{
			name: "Apple profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
		},
		{
			name: "Pear profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Pear")}},
		},
		{
			name: "all records",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
		},
		{
			name: "Cherry single",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
		},
		{
			name: "Height <= 12",
			crit: [][]Value{{StringVal("Height")}, {StringVal("<=12")}},
		},
	}

	db := standardDB()
	for _, sc := range scenarios {
		t.Run(sc.name, func(t *testing.T) {
			resolver := makeDBResolver(db, 1, 1, sc.crit, 7, 1)
			critRange := dbRange(7, 1, len(sc.crit[0]), len(sc.crit))

			varpFormula := `DVARP(A1:E7,"Profit",` + critRange + `)`
			stdevpFormula := `DSTDEVP(A1:E7,"Profit",` + critRange + `)`

			cfV := evalCompile(t, varpFormula)
			gotV, err := Eval(cfV, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DVARP: %v", err)
			}

			cfP := evalCompile(t, stdevpFormula)
			gotP, err := Eval(cfP, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DSTDEVP: %v", err)
			}

			if gotV.Type != ValueNumber || gotP.Type != ValueNumber {
				return
			}

			stdevpSquared := gotP.Num * gotP.Num
			diff := gotV.Num - stdevpSquared
			if diff > 1e-6 || diff < -1e-6 {
				t.Errorf("DVARP (%g) != DSTDEVP^2 (%g), diff=%g", gotV.Num, stdevpSquared, diff)
			}
		})
	}
}

// TestDSTDEVP_LessOrEqual_DSTDEV verifies that DSTDEVP <= DSTDEV for the same data.
// Population stdev (n denominator) is always <= sample stdev (n-1 denominator).
func TestDSTDEVP_LessOrEqual_DSTDEV(t *testing.T) {
	scenarios := []struct {
		name string
		crit [][]Value
	}{
		{
			name: "Apple profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
		},
		{
			name: "Pear profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Pear")}},
		},
		{
			name: "all records profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
		},
		{
			name: "Height > 10",
			crit: [][]Value{{StringVal("Height")}, {StringVal(">10")}},
		},
	}

	db := standardDB()
	for _, sc := range scenarios {
		t.Run(sc.name, func(t *testing.T) {
			resolver := makeDBResolver(db, 1, 1, sc.crit, 7, 1)
			critRange := dbRange(7, 1, len(sc.crit[0]), len(sc.crit))

			stdevFormula := `DSTDEV(A1:E7,"Profit",` + critRange + `)`
			stdevpFormula := `DSTDEVP(A1:E7,"Profit",` + critRange + `)`

			cfS := evalCompile(t, stdevFormula)
			gotS, err := Eval(cfS, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DSTDEV: %v", err)
			}

			cfP := evalCompile(t, stdevpFormula)
			gotP, err := Eval(cfP, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DSTDEVP: %v", err)
			}

			if gotS.Type != ValueNumber || gotP.Type != ValueNumber {
				return
			}

			if gotP.Num > gotS.Num+1e-9 {
				t.Errorf("DSTDEVP (%g) > DSTDEV (%g), expected DSTDEVP <= DSTDEV", gotP.Num, gotS.Num)
			}
		})
	}
}

// TestDVARP_LessOrEqual_DVAR verifies that DVARP <= DVAR for the same data.
func TestDVARP_LessOrEqual_DVAR(t *testing.T) {
	scenarios := []struct {
		name string
		crit [][]Value
	}{
		{
			name: "Apple profit",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
		},
		{
			name: "all records",
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
		},
		{
			name: "Height < 15",
			crit: [][]Value{{StringVal("Height")}, {StringVal("<15")}},
		},
	}

	db := standardDB()
	for _, sc := range scenarios {
		t.Run(sc.name, func(t *testing.T) {
			resolver := makeDBResolver(db, 1, 1, sc.crit, 7, 1)
			critRange := dbRange(7, 1, len(sc.crit[0]), len(sc.crit))

			varFormula := `DVAR(A1:E7,"Profit",` + critRange + `)`
			varpFormula := `DVARP(A1:E7,"Profit",` + critRange + `)`

			cfV := evalCompile(t, varFormula)
			gotV, err := Eval(cfV, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DVAR: %v", err)
			}

			cfP := evalCompile(t, varpFormula)
			gotP, err := Eval(cfP, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DVARP: %v", err)
			}

			if gotV.Type != ValueNumber || gotP.Type != ValueNumber {
				return
			}

			if gotP.Num > gotV.Num+1e-9 {
				t.Errorf("DVARP (%g) > DVAR (%g), expected DVARP <= DVAR", gotP.Num, gotV.Num)
			}
		})
	}
}

// TestDAVERAGE_DVAR_Consistency verifies that for data with known mean and variance,
// the DAVERAGE and DVAR results are consistent.
func TestDAVERAGE_DVAR_Consistency(t *testing.T) {
	// Database: values 2, 4, 4, 4, 5, 5, 7, 9
	// Known: mean=5, population variance=4, sample variance=4.571...
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(2)},
		{StringVal("B"), NumberVal(4)},
		{StringVal("C"), NumberVal(4)},
		{StringVal("D"), NumberVal(4)},
		{StringVal("E"), NumberVal(5)},
		{StringVal("F"), NumberVal(5)},
		{StringVal("G"), NumberVal(7)},
		{StringVal("H"), NumberVal(9)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	// Check DAVERAGE = 5
	avgFormula := `DAVERAGE(A1:B9,"Value",G1:G2)`
	cfA := evalCompile(t, avgFormula)
	gotA, err := Eval(cfA, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DAVERAGE: %v", err)
	}
	if gotA.Type != ValueNumber {
		t.Fatalf("DAVERAGE type = %v, want number", gotA.Type)
	}
	if diff := gotA.Num - 5.0; diff > 1e-9 || diff < -1e-9 {
		t.Errorf("DAVERAGE = %g, want 5", gotA.Num)
	}

	// Check DVARP = 4
	varpFormula := `DVARP(A1:B9,"Value",G1:G2)`
	cfVP := evalCompile(t, varpFormula)
	gotVP, err := Eval(cfVP, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DVARP: %v", err)
	}
	if gotVP.Type != ValueNumber {
		t.Fatalf("DVARP type = %v, want number", gotVP.Type)
	}
	if diff := gotVP.Num - 4.0; diff > 1e-9 || diff < -1e-9 {
		t.Errorf("DVARP = %g, want 4", gotVP.Num)
	}

	// Check DVAR = 32/7 = 4.571428...
	varFormula := `DVAR(A1:B9,"Value",G1:G2)`
	cfV := evalCompile(t, varFormula)
	gotV, err := Eval(cfV, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DVAR: %v", err)
	}
	if gotV.Type != ValueNumber {
		t.Fatalf("DVAR type = %v, want number", gotV.Type)
	}
	expected := 32.0 / 7.0
	if diff := gotV.Num - expected; diff > 1e-9 || diff < -1e-9 {
		t.Errorf("DVAR = %g, want %g", gotV.Num, expected)
	}

	// Check DSTDEVP = 2
	stdevpFormula := `DSTDEVP(A1:B9,"Value",G1:G2)`
	cfSP := evalCompile(t, stdevpFormula)
	gotSP, err := Eval(cfSP, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSTDEVP: %v", err)
	}
	if gotSP.Type != ValueNumber {
		t.Fatalf("DSTDEVP type = %v, want number", gotSP.Type)
	}
	if diff := gotSP.Num - 2.0; diff > 1e-9 || diff < -1e-9 {
		t.Errorf("DSTDEVP = %g, want 2", gotSP.Num)
	}
}

// TestDSTDEV_DSTDEVP_SingleMatch_Diverge tests that DSTDEV returns #DIV/0!
// for a single match while DSTDEVP returns 0.
func TestDSTDEV_DSTDEVP_SingleMatch_Diverge(t *testing.T) {
	db := standardDB()
	crit := [][]Value{
		{StringVal("Tree")},
		{StringVal("Cherry")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	// DSTDEV should return #DIV/0! (sample needs n>=2)
	stdevFormula := `DSTDEV(A1:E7,"Profit",G1:G2)`
	cfS := evalCompile(t, stdevFormula)
	gotS, err := Eval(cfS, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSTDEV: %v", err)
	}
	if gotS.Type != ValueError || gotS.Err != ErrValDIV0 {
		t.Errorf("DSTDEV single match = %+v, want #DIV/0!", gotS)
	}

	// DSTDEVP should return 0
	stdevpFormula := `DSTDEVP(A1:E7,"Profit",G1:G2)`
	cfP := evalCompile(t, stdevpFormula)
	gotP, err := Eval(cfP, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DSTDEVP: %v", err)
	}
	if gotP.Type != ValueNumber || gotP.Num != 0 {
		t.Errorf("DSTDEVP single match = %+v, want 0", gotP)
	}
}

// TestDVAR_DVARP_SingleMatch_Diverge tests that DVAR returns #DIV/0!
// for a single match while DVARP returns 0.
func TestDVAR_DVARP_SingleMatch_Diverge(t *testing.T) {
	db := standardDB()
	crit := [][]Value{
		{StringVal("Tree")},
		{StringVal("Cherry")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	// DVAR should return #DIV/0!
	varFormula := `DVAR(A1:E7,"Profit",G1:G2)`
	cfV := evalCompile(t, varFormula)
	gotV, err := Eval(cfV, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DVAR: %v", err)
	}
	if gotV.Type != ValueError || gotV.Err != ErrValDIV0 {
		t.Errorf("DVAR single match = %+v, want #DIV/0!", gotV)
	}

	// DVARP should return 0
	varpFormula := `DVARP(A1:E7,"Profit",G1:G2)`
	cfP := evalCompile(t, varpFormula)
	gotP, err := Eval(cfP, resolver, nil)
	if err != nil {
		t.Fatalf("Eval DVARP: %v", err)
	}
	if gotP.Type != ValueNumber || gotP.Num != 0 {
		t.Errorf("DVARP single match = %+v, want 0", gotP)
	}
}

// TestDAVERAGE_ZeroDivide_Vs_DSTDEV tests that both functions return #DIV/0!
// when there are no numeric matching values.
func TestDAVERAGE_ZeroDivide_Vs_DSTDEV(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), StringVal("text")},
		{StringVal("B"), StringVal("text2")},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}
	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	for _, fn := range []string{"DAVERAGE", "DSTDEV", "DSTDEVP", "DVAR", "DVARP"} {
		t.Run(fn, func(t *testing.T) {
			formula := fn + `(A1:B3,"Value",G1:G2)`
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}
			if got.Type != ValueError || got.Err != ErrValDIV0 {
				t.Errorf("%s with all text = %+v, want #DIV/0!", fn, got)
			}
		})
	}
}

// TestDStatFunctions_WrongArgCount_All tests wrong arg count for all five stat functions.
func TestDStatFunctions_WrongArgCount_All(t *testing.T) {
	type fnType func([]Value) (Value, error)
	funcs := []struct {
		name string
		fn   fnType
	}{
		{"DAVERAGE", fnDAverage},
		{"DSTDEV", fnDStdev},
		{"DSTDEVP", fnDStdevP},
		{"DVAR", fnDVar},
		{"DVARP", fnDVarP},
	}
	argSets := []struct {
		label string
		args  []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}

	for _, fn := range funcs {
		for _, as := range argSets {
			t.Run(fn.name+"/"+as.label, func(t *testing.T) {
				result, err := fn.fn(as.args)
				if err != nil {
					t.Fatalf("%s error: %v", fn.name, err)
				}
				if result.Type != ValueError || result.Err != ErrValVALUE {
					t.Errorf("%s(%s) = %+v, want #VALUE!", fn.name, as.label, result)
				}
			})
		}
	}
}

// ---------------------------------------------------------------------------
// Error propagation tests for new D-functions
// ---------------------------------------------------------------------------

func TestDFunctions_ErrorPropagation(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	funcs := []string{"DAVERAGE", "DCOUNT", "DCOUNTA", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP"}
	for _, fn := range funcs {
		t.Run(fn, func(t *testing.T) {
			formula := fn + `(A1:B4,"Value",G1:G2)`
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}
			if got.Type != ValueError || got.Err != ErrValDIV0 {
				t.Errorf("%s with error cell = %+v, want #DIV/0!", fn, got)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Cross-check tests: DMAX >= DMIN and DCOUNT <= DCOUNTA
// ---------------------------------------------------------------------------

func TestDMAX_GE_DMIN_CrossCheck(t *testing.T) {
	// For the same database and criteria, DMAX should always be >= DMIN.
	testCases := []struct {
		name string
		db   [][]Value
		crit [][]Value
		field string
	}{
		{
			name: "standard DB all records",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
			field: `"Profit"`,
		},
		{
			name: "standard DB Apple only",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field: `"Profit"`,
		},
		{
			name: "standard DB Pear only",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Pear")}},
			field: `"Height"`,
		},
		{
			name: "standard DB Height>10",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Height")}, {StringVal(">10")}},
			field: `"Yield"`,
		},
		{
			name: "negative values",
			db: [][]Value{
				{StringVal("Name"), StringVal("Val")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(-20)},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Val"`,
		},
		{
			name: "mixed positive and negative",
			db: [][]Value{
				{StringVal("Name"), StringVal("Val")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(0)},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Val"`,
		},
		{
			name: "single record",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field: `"Profit"`,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			dbRows := len(tc.db)
			dbCols := len(tc.db[0])
			critRows := len(tc.crit)
			critCols := len(tc.crit[0])

			resolver := makeDBResolver(tc.db, 1, 1, tc.crit, 7, 1)

			maxFormula := "DMAX(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tc.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"
			minFormula := "DMIN(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tc.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cfMax := evalCompile(t, maxFormula)
			gotMax, err := Eval(cfMax, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DMAX: %v", err)
			}

			cfMin := evalCompile(t, minFormula)
			gotMin, err := Eval(cfMin, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DMIN: %v", err)
			}

			if gotMax.Type != ValueNumber || gotMin.Type != ValueNumber {
				t.Fatalf("Expected numbers: DMAX=%+v, DMIN=%+v", gotMax, gotMin)
			}

			if gotMax.Num < gotMin.Num {
				t.Errorf("DMAX(%g) < DMIN(%g), expected DMAX >= DMIN", gotMax.Num, gotMin.Num)
			}
		})
	}
}

func TestDCOUNT_LE_DCOUNTA_CrossCheck(t *testing.T) {
	// For the same database and criteria, DCOUNT should always be <= DCOUNTA.
	testCases := []struct {
		name string
		db   [][]Value
		crit [][]Value
		field string
	}{
		{
			name: "standard DB all records - numeric column",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
			field: `"Profit"`,
		},
		{
			name: "standard DB all records - text column",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("")}},
			field: `"Tree"`,
		},
		{
			name: "mixed types column",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Value"`,
		},
		{
			name: "all text",
			db: [][]Value{
				{StringVal("Name"), StringVal("Val")},
				{StringVal("A"), StringVal("x")},
				{StringVal("B"), StringVal("y")},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Val"`,
		},
		{
			name: "all empty",
			db: [][]Value{
				{StringVal("Name"), StringVal("Val")},
				{StringVal("A"), EmptyVal()},
				{StringVal("B"), EmptyVal()},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Val"`,
		},
		{
			name: "all booleans",
			db: [][]Value{
				{StringVal("Name"), StringVal("Flag")},
				{StringVal("A"), BoolVal(true)},
				{StringVal("B"), BoolVal(false)},
				{StringVal("C"), BoolVal(true)},
			},
			crit: [][]Value{{StringVal("Name")}, {StringVal("")}},
			field: `"Flag"`,
		},
		{
			name: "standard DB Apple - numeric field",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field: `"Profit"`,
		},
		{
			name: "no matches",
			db:   standardDB(),
			crit: [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field: `"Profit"`,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			dbRows := len(tc.db)
			dbCols := len(tc.db[0])
			critRows := len(tc.crit)
			critCols := len(tc.crit[0])

			resolver := makeDBResolver(tc.db, 1, 1, tc.crit, 7, 1)

			dcountFormula := "DCOUNT(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tc.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"
			dcountaFormula := "DCOUNTA(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tc.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cfCount := evalCompile(t, dcountFormula)
			gotCount, err := Eval(cfCount, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DCOUNT: %v", err)
			}

			cfCountA := evalCompile(t, dcountaFormula)
			gotCountA, err := Eval(cfCountA, resolver, nil)
			if err != nil {
				t.Fatalf("Eval DCOUNTA: %v", err)
			}

			if gotCount.Type != ValueNumber || gotCountA.Type != ValueNumber {
				t.Fatalf("Expected numbers: DCOUNT=%+v, DCOUNTA=%+v", gotCount, gotCountA)
			}

			if gotCount.Num > gotCountA.Num {
				t.Errorf("DCOUNT(%g) > DCOUNTA(%g), expected DCOUNT <= DCOUNTA", gotCount.Num, gotCountA.Num)
			}
		})
	}
}
