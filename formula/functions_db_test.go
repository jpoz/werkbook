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
			name:     "var single value returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "var no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
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
		{
			name:     "varp no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
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
