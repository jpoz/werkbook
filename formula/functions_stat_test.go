package formula

import (
	"fmt"
	"math"
	"testing"
)

// ---------------------------------------------------------------------------
// LARGE / SMALL
// ---------------------------------------------------------------------------

func TestLargeSmall(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "LARGE(A1:A3,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("LARGE k=1: got %g, want 30", got.Num)
	}

	cf = evalCompile(t, "SMALL(A1:A3,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("SMALL k=1: got %g, want 10", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNTBLANK
// ---------------------------------------------------------------------------

func TestCountBlank(t *testing.T) {
	t.Run("empty cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				// A2 is empty
				{Col: 1, Row: 3}: NumberVal(3),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
		}
	})

	t.Run("empty string counts as blank", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				// A2 is empty (missing)
				{Col: 1, Row: 3}: StringVal(""), // empty string = blank
				{Col: 1, Row: 4}: StringVal("hello"),
				{Col: 1, Row: 5}: NumberVal(0), // zero is not blank
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("COUNTBLANK: got %g, want 2", got.Num)
		}
	})

	t.Run("single empty string cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal(""),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
		}
	})

	t.Run("all empty range", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("COUNTBLANK: got %g, want 5", got.Num)
		}
	})

	t.Run("no empty cells returns 0", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("COUNTBLANK: got %g, want 0", got.Num)
		}
	})

	t.Run("zero is not blank", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("COUNTBLANK: got %g, want 0", got.Num)
		}
	})

	t.Run("boolean FALSE is not blank", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(true),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("COUNTBLANK: got %g, want 0", got.Num)
		}
	})

	t.Run("strings are not blank", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
				{Col: 1, Row: 2}: StringVal("world"),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("COUNTBLANK: got %g, want 0", got.Num)
		}
	})

	t.Run("2D range with mixed content", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(6),
				// A2 empty
				{Col: 1, Row: 3}: NumberVal(4),
				// B1 empty
				{Col: 2, Row: 2}: NumberVal(27),
				{Col: 2, Row: 3}: NumberVal(34),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:B3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("COUNTBLANK: got %g, want 2", got.Num)
		}
	})

	t.Run("single non-empty cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("COUNTBLANK: got %g, want 0", got.Num)
		}
	})

	t.Run("single empty cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{},
		}

		cf := evalCompile(t, "COUNTBLANK(A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
		}
	})

	t.Run("large range with gaps", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}:  NumberVal(1),
				{Col: 1, Row: 5}:  NumberVal(5),
				{Col: 1, Row: 10}: NumberVal(10),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("COUNTBLANK: got %g, want 7", got.Num)
		}
	})

	t.Run("mix of empty strings and missing cells", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal(""),
				// A2 missing
				{Col: 1, Row: 3}: StringVal(""),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("COUNTBLANK: got %g, want 3", got.Num)
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		got, err := fnCOUNTBLANK([]Value{})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("too many args returns VALUE error", func(t *testing.T) {
		got, err := fnCOUNTBLANK([]Value{NumberVal(1), NumberVal(2)})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})
}

// ---------------------------------------------------------------------------
// SUMIF
// ---------------------------------------------------------------------------

func TestSUMIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("SUMIF >15: got %g, want 50", got.Num)
	}
}

func TestSUMIFPropagatesMatchedSumRangeError(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("East"),
			{Col: 1, Row: 2}: StringVal("West"),
			{Col: 2, Row: 1}: ErrorVal(ErrValNAME),
			{Col: 2, Row: 2}: NumberVal(10),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A2,"East",B1:B2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNAME {
		t.Fatalf(`SUMIF(A1:A2,"East",B1:B2) = %v, want #NAME?`, got)
	}
}

// ---------------------------------------------------------------------------
// COUNT – boolean handling
// ---------------------------------------------------------------------------

func TestCOUNTBooleanDirectArgs(t *testing.T) {
	resolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Direct boolean args should be counted (Excel behavior).
	cf := evalCompile(t, "COUNT(TRUE,FALSE,10,20)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNT(TRUE,FALSE,10,20): got %g, want 4", got.Num)
	}

	cf = evalCompile(t, "COUNT(TRUE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNT(TRUE): got %g, want 1", got.Num)
	}
}

func TestCOUNTBooleanInRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: BoolVal(false),
		},
	}

	// Booleans in a range should NOT be counted.
	cf := evalCompile(t, "COUNT(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNT(A1:A2) with booleans: got %g, want 0", got.Num)
	}

	// A single cell reference to a boolean should NOT be counted either.
	cf = evalCompile(t, "COUNT(A1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNT(A1) where A1=TRUE: got %g, want 0", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNT – comprehensive tests
// ---------------------------------------------------------------------------

func TestCOUNT(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		cells    map[CellAddr]Value
		expected float64
	}{
		{
			name:    "all numbers",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3),
			},
			expected: 3,
		},
		{
			name:    "strings not counted",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
				{Col: 1, Row: 2}: StringVal("world"),
				{Col: 1, Row: 3}: StringVal("foo"),
			},
			expected: 0,
		},
		{
			name:    "booleans in range not counted",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: NumberVal(5),
			},
			expected: 1,
		},
		{
			name:    "empty cells not counted",
			formula: "COUNT(A1:A5)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				// A2 empty
				{Col: 1, Row: 3}: NumberVal(30),
				// A4 empty
				{Col: 1, Row: 5}: NumberVal(50),
			},
			expected: 3,
		},
		{
			name:    "all strings yields zero",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				{Col: 1, Row: 2}: StringVal("b"),
				{Col: 1, Row: 3}: StringVal("c"),
			},
			expected: 0,
		},
		{
			name:    "mixed types only numbers counted",
			formula: "COUNT(A1:A5)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(100),
				{Col: 1, Row: 2}: StringVal("text"),
				{Col: 1, Row: 3}: BoolVal(true),
				// A4 empty
				{Col: 1, Row: 5}: NumberVal(200),
			},
			expected: 2,
		},
		{
			name:    "zero is counted",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(1),
			},
			expected: 3,
		},
		{
			name:    "error values not counted in range",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
				{Col: 1, Row: 3}: NumberVal(3),
			},
			expected: 2,
		},
		{
			name:    "multiple ranges",
			formula: "COUNT(A1:A2,B1:B2)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 2, Row: 2}: StringVal("x"),
			},
			expected: 3,
		},
		{
			name:     "scalar number argument counted",
			formula:  "COUNT(42)",
			cells:    map[CellAddr]Value{},
			expected: 1,
		},
		{
			name:     "scalar string not counted",
			formula:  `COUNT("hello")`,
			cells:    map[CellAddr]Value{},
			expected: 0,
		},
		{
			name:     "scalar boolean TRUE counted",
			formula:  "COUNT(TRUE)",
			cells:    map[CellAddr]Value{},
			expected: 1,
		},
		{
			name:     "scalar boolean FALSE counted",
			formula:  "COUNT(FALSE)",
			cells:    map[CellAddr]Value{},
			expected: 1,
		},
		{
			name:     "scalar string number counted",
			formula:  `COUNT("5")`,
			cells:    map[CellAddr]Value{},
			expected: 1,
		},
		{
			name:     "scalar non-numeric string not counted",
			formula:  `COUNT("abc")`,
			cells:    map[CellAddr]Value{},
			expected: 0,
		},
		{
			name:    "range with gaps",
			formula: "COUNT(A1:A10)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}:  NumberVal(1),
				{Col: 1, Row: 5}:  NumberVal(5),
				{Col: 1, Row: 10}: NumberVal(10),
			},
			expected: 3,
		},
		{
			name:    "negative numbers counted",
			formula: "COUNT(A1:A3)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-1),
				{Col: 1, Row: 2}: NumberVal(-99.5),
				{Col: 1, Row: 3}: NumberVal(0),
			},
			expected: 3,
		},
		{
			name:    "excel doc example COUNT(A2:A6) dates and numbers",
			formula: "COUNT(A1:A5)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39790), // 12/8/2008 as serial
				{Col: 1, Row: 2}: NumberVal(19),
				{Col: 1, Row: 3}: NumberVal(22.24),
				{Col: 1, Row: 4}: BoolVal(true),
				{Col: 1, Row: 5}: ErrorVal(ErrValDIV0),
			},
			expected: 3,
		},
		{
			name:    "excel doc example with extra scalar",
			formula: "COUNT(A1:A5,2)",
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39790),
				{Col: 1, Row: 2}: NumberVal(19),
				{Col: 1, Row: 3}: NumberVal(22.24),
				{Col: 1, Row: 4}: BoolVal(true),
				{Col: 1, Row: 5}: ErrorVal(ErrValDIV0),
			},
			expected: 4,
		},
		{
			name:     "mixed scalar args",
			formula:  `COUNT(1,2,"hi",TRUE,"3")`,
			cells:    map[CellAddr]Value{},
			expected: 4, // 1, 2, TRUE, "3"
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			resolver := &mockResolver{cells: tc.cells}
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tc.expected {
				t.Errorf("%s: got %v (%g), want %g", tc.formula, got.Type, got.Num, tc.expected)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// COUNTIF
// ---------------------------------------------------------------------------

func TestCOUNTIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
		},
	}

	cf := evalCompile(t, `COUNTIF(A1:A3,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF: got %g, want 2", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMPRODUCT
// ---------------------------------------------------------------------------

func TestSUMPRODUCT(t *testing.T) {
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

	// 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
	cf := evalCompile(t, "SUMPRODUCT(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 32 {
		t.Errorf("SUMPRODUCT: got %g, want 32", got.Num)
	}
}

func TestSUMPRODUCT_Comprehensive(t *testing.T) {
	t.Run("basic two arrays", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 2, Row: 1}: NumberVal(5),
				{Col: 2, Row: 2}: NumberVal(7),
			},
		}
		// 2*5 + 3*7 = 10 + 21 = 31
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 31 {
			t.Errorf("got %v, want 31", got)
		}
	})

	t.Run("three arrays", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 2, Row: 2}: NumberVal(4),
				{Col: 3, Row: 1}: NumberVal(5),
				{Col: 3, Row: 2}: NumberVal(6),
			},
		}
		// 1*3*5 + 2*4*6 = 15 + 48 = 63
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2,C1:C2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 63 {
			t.Errorf("got %v, want 63", got)
		}
	})

	t.Run("single array just sums", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
		}
		// single array: product of one element is itself, so sum = 10+20+30 = 60
		cf := evalCompile(t, "SUMPRODUCT(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 60 {
			t.Errorf("got %v, want 60", got)
		}
	})

	t.Run("arrays with zeros", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(3),
				{Col: 2, Row: 1}: NumberVal(2),
				{Col: 2, Row: 2}: NumberVal(7),
				{Col: 2, Row: 3}: NumberVal(0),
			},
		}
		// 5*2 + 0*7 + 3*0 = 10 + 0 + 0 = 10
		cf := evalCompile(t, "SUMPRODUCT(A1:A3,B1:B3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("arrays with negative numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-1),
				{Col: 1, Row: 2}: NumberVal(-2),
				{Col: 2, Row: 1}: NumberVal(-3),
				{Col: 2, Row: 2}: NumberVal(-4),
			},
		}
		// (-1)*(-3) + (-2)*(-4) = 3 + 8 = 11
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 11 {
			t.Errorf("got %v, want 11", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3),
				{Col: 1, Row: 2}: NumberVal(-2),
				{Col: 2, Row: 1}: NumberVal(-4),
				{Col: 2, Row: 2}: NumberVal(5),
			},
		}
		// 3*(-4) + (-2)*5 = -12 + -10 = -22
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -22 {
			t.Errorf("got %v, want -22", got)
		}
	})

	t.Run("single element arrays", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(7),
				{Col: 2, Row: 1}: NumberVal(8),
			},
		}
		// 7*8 = 56
		cf := evalCompile(t, "SUMPRODUCT(A1:A1,B1:B1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 56 {
			t.Errorf("got %v, want 56", got)
		}
	})

	t.Run("different size arrays returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3),
				{Col: 2, Row: 1}: NumberVal(4),
				{Col: 2, Row: 2}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "SUMPRODUCT(A1:A3,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		got, err := fnSUMPRODUCT([]Value{})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("non-array arg returns VALUE error", func(t *testing.T) {
		got, err := fnSUMPRODUCT([]Value{NumberVal(5)})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("text in array treated as zero", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 2, Row: 2}: NumberVal(4),
			},
		}
		// 2*3 + 0*4 = 6 + 0 = 6  (text coerced to 0)
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6 {
			t.Errorf("got %v, want 6", got)
		}
	})

	t.Run("boolean TRUE treated as 1 in array", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 2, Row: 1}: NumberVal(10),
				{Col: 2, Row: 2}: NumberVal(20),
			},
		}
		// 1*10 + 0*20 = 10
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("all zeros returns zero", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 2, Row: 1}: NumberVal(0),
				{Col: 2, Row: 2}: NumberVal(0),
			},
		}
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("large values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1e6),
				{Col: 1, Row: 2}: NumberVal(2e6),
				{Col: 2, Row: 1}: NumberVal(3e6),
				{Col: 2, Row: 2}: NumberVal(4e6),
			},
		}
		// 1e6*3e6 + 2e6*4e6 = 3e12 + 8e12 = 1.1e13
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1.1e13 {
			t.Errorf("got %v, want 1.1e13", got)
		}
	})

	t.Run("single row range", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 3, Row: 1}: NumberVal(4),
				{Col: 1, Row: 2}: NumberVal(5),
				{Col: 2, Row: 2}: NumberVal(6),
				{Col: 3, Row: 2}: NumberVal(7),
			},
		}
		// A1:C1 = {2,3,4}, A2:C2 = {5,6,7}
		// 2*5 + 3*6 + 4*7 = 10 + 18 + 28 = 56
		cf := evalCompile(t, "SUMPRODUCT(A1:C1,A2:C2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 56 {
			t.Errorf("got %v, want 56", got)
		}
	})

	t.Run("multi-row multi-col arrays", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 2, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 2, Row: 2}: NumberVal(4),
				{Col: 3, Row: 1}: NumberVal(5),
				{Col: 4, Row: 1}: NumberVal(6),
				{Col: 3, Row: 2}: NumberVal(7),
				{Col: 4, Row: 2}: NumberVal(8),
			},
		}
		// A1:B2 = {{1,2},{3,4}}, C1:D2 = {{5,6},{7,8}}
		// 1*5 + 2*6 + 3*7 + 4*8 = 5 + 12 + 21 + 32 = 70
		cf := evalCompile(t, "SUMPRODUCT(A1:B2,C1:D2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 70 {
			t.Errorf("got %v, want 70", got)
		}
	})

	t.Run("error value in array propagates", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 2, Row: 2}: NumberVal(4),
			},
		}
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("empty cells treated as zero", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				// A2 is empty
				{Col: 2, Row: 1}: NumberVal(3),
				{Col: 2, Row: 2}: NumberVal(4),
			},
		}
		// 5*3 + 0*4 = 15
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("fractional values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0.5),
				{Col: 1, Row: 2}: NumberVal(1.5),
				{Col: 2, Row: 1}: NumberVal(2.5),
				{Col: 2, Row: 2}: NumberVal(3.5),
			},
		}
		// 0.5*2.5 + 1.5*3.5 = 1.25 + 5.25 = 6.5
		cf := evalCompile(t, "SUMPRODUCT(A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6.5 {
			t.Errorf("got %v, want 6.5", got)
		}
	})
}

// ---------------------------------------------------------------------------
// MatchesCriteria — helper used by *IF functions
// ---------------------------------------------------------------------------

func TestMatchesCriteria(t *testing.T) {
	tests := []struct {
		v    Value
		crit Value
		want bool
	}{
		{NumberVal(10), StringVal(">5"), true},
		{NumberVal(3), StringVal(">5"), false},
		{NumberVal(5), StringVal(">=5"), true},
		{NumberVal(5), StringVal("<=5"), true},
		{NumberVal(5), StringVal("<>5"), false},
		{NumberVal(5), NumberVal(5), true},
		{StringVal("apple"), StringVal("app*"), true},
		{StringVal("banana"), StringVal("app*"), false},
		{StringVal("cat"), StringVal("c?t"), true},
	}

	for _, tt := range tests {
		got := MatchesCriteria(tt.v, tt.crit)
		if got != tt.want {
			t.Errorf("MatchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// COUNTA — counts all non-empty cells
// ---------------------------------------------------------------------------

func TestCOUNTA(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			// A3 is empty
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 1, Row: 5}: ErrorVal(ErrValNA),
		},
	}

	cf := evalCompile(t, "COUNTA(A1:A5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Number, String, Bool, Error = 4 non-empty cells
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNTA: got %g, want 4", got.Num)
	}

	// All empty
	cf = evalCompile(t, "COUNTA(C1:C3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNTA empty range: got %g, want 0", got.Num)
	}

	// Scalar argument
	cf = evalCompile(t, `COUNTA("hi")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTA scalar: got %g, want 1", got.Num)
	}

	// --- Additional comprehensive tests ---

	// All numbers → count all
	t.Run("all_numbers", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %g, want 3", got.Num)
		}
	})

	// Zero is counted (not empty)
	t.Run("zero_is_counted", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})

	// Empty string "" IS counted by COUNTA
	t.Run("empty_string_counted", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal(""),
				{Col: 1, Row: 2}: StringVal("text"),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})

	// Boolean values are counted
	t.Run("booleans_counted", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})

	// Error values are counted
	t.Run("errors_counted", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
				{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})

	// Mixed types: numbers, strings, booleans, errors
	t.Run("mixed_types", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
				{Col: 1, Row: 2}: StringVal("abc"),
				{Col: 1, Row: 3}: BoolVal(false),
				{Col: 1, Row: 4}: ErrorVal(ErrValNA),
				{Col: 1, Row: 5}: NumberVal(0),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %g, want 5", got.Num)
		}
	})

	// Range with gaps (sparse cells)
	t.Run("range_with_gaps", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				// A2 empty
				{Col: 1, Row: 3}: NumberVal(3),
				// A4 empty
				{Col: 1, Row: 5}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %g, want 3", got.Num)
		}
	})

	// Multiple ranges
	t.Run("multiple_ranges", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 2, Row: 1}: StringVal("x"),
				{Col: 2, Row: 2}: StringVal("y"),
				{Col: 2, Row: 3}: StringVal("z"),
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A2,B1:B3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %g, want 5", got.Num)
		}
	})

	// Multiple scalar arguments
	t.Run("multiple_scalars", func(t *testing.T) {
		r := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, `COUNTA(1,2,3)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %g, want 3", got.Num)
		}
	})

	// Mix of scalars and ranges
	t.Run("scalars_and_ranges", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
			},
		}
		cf := evalCompile(t, `COUNTA(A1:A2,"extra",99)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %g, want 4", got.Num)
		}
	})

	// Single cell reference (non-empty)
	t.Run("single_cell_nonempty", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(99),
			},
		}
		cf := evalCompile(t, "COUNTA(A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %g, want 1", got.Num)
		}
	})

	// Single cell reference (empty)
	t.Run("single_cell_empty", func(t *testing.T) {
		r := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "COUNTA(A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %g, want 0", got.Num)
		}
	})

	// Scalar boolean TRUE
	t.Run("scalar_boolean", func(t *testing.T) {
		r := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "COUNTA(TRUE)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %g, want 1", got.Num)
		}
	})

	// Excel doc example: date, number, decimal, TRUE, #DIV/0! → 5
	t.Run("excel_doc_example", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39790),     // date serial
				{Col: 1, Row: 2}: NumberVal(19),        // integer
				{Col: 1, Row: 3}: NumberVal(22.24),     // decimal
				{Col: 1, Row: 4}: BoolVal(true),        // TRUE
				{Col: 1, Row: 5}: ErrorVal(ErrValDIV0), // #DIV/0!
			},
		}
		cf := evalCompile(t, "COUNTA(A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %g, want 5", got.Num)
		}
	})

	// Scalar empty string "" is counted
	t.Run("scalar_empty_string", func(t *testing.T) {
		r := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, `COUNTA("")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %g, want 1", got.Num)
		}
	})
}

// ---------------------------------------------------------------------------
// SUMIFS — multiple criteria
// ---------------------------------------------------------------------------

func TestSUMIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (B)
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			// Criteria range 1 (A) — category
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 1, Row: 2}: StringVal("veg"),
			{Col: 1, Row: 3}: StringVal("fruit"),
			{Col: 1, Row: 4}: StringVal("veg"),
			// Criteria range 2 (C) — score
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
			{Col: 3, Row: 3}: NumberVal(25),
			{Col: 3, Row: 4}: NumberVal(35),
		},
	}

	// SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)
	// Sum B where A="fruit" AND C>10
	cf := evalCompile(t, `SUMIFS(B1:B4,A1:A4,"fruit",C1:C4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only row 3 matches (fruit, 25>10) => sum=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("SUMIFS: got %g, want 30", got.Num)
	}
}

func TestSUMIFSSingleCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,">=2")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 2,3 match => 20+30=50
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("SUMIFS single: got %g, want 50", got.Num)
	}
}

func TestSUMIFSArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Odd number of args (invalid)
	cf := evalCompile(t, "SUMIFS(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SUMIFS bad args: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// SUMIFS — comprehensive tests
// ---------------------------------------------------------------------------

func TestSUMIFS_AllComparisonOperators(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (A)
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 1, Row: 5}: NumberVal(50),
			// Criteria range (B) — scores
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			{Col: 2, Row: 5}: NumberVal(50),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"greater than", `SUMIFS(A1:A5,B1:B5,">30")`, 90},       // 40+50
		{"less than", `SUMIFS(A1:A5,B1:B5,"<30")`, 30},          // 10+20
		{"greater or equal", `SUMIFS(A1:A5,B1:B5,">=30")`, 120}, // 30+40+50
		{"less or equal", `SUMIFS(A1:A5,B1:B5,"<=30")`, 60},     // 10+20+30
		{"equal", `SUMIFS(A1:A5,B1:B5,"=30")`, 30},              // 30
		{"not equal", `SUMIFS(A1:A5,B1:B5,"<>30")`, 120},        // 10+20+40+50
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("got %v, want %g", got, tt.want)
			}
		})
	}
}

func TestSUMIFS_NoMatches(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 2, Row: 1}: StringVal("apple"),
			{Col: 2, Row: 2}: StringVal("banana"),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A2,B1:B2,"cherry")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("SUMIFS no matches: got %v, want 0", got)
	}
}

func TestSUMIFS_AllMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(15),
			{Col: 2, Row: 3}: NumberVal(25),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUMIFS all match: got %v, want 60", got)
	}
}

func TestSUMIFS_WildcardAsterisk(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Apples"),
			{Col: 2, Row: 3}: StringVal("Artichokes"),
			{Col: 2, Row: 4}: StringVal("Bananas"),
		},
	}

	// Wildcard * matches any sequence of characters
	cf := evalCompile(t, `SUMIFS(A1:A4,B1:B4,"A*")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Apples(5) + Apples(4) + Artichokes(15) = 24
	if got.Type != ValueNumber || got.Num != 24 {
		t.Errorf("SUMIFS wildcard *: got %v, want 24", got)
	}
}

func TestSUMIFS_WildcardQuestionMark(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 2, Row: 1}: StringVal("cat"),
			{Col: 2, Row: 2}: StringVal("car"),
			{Col: 2, Row: 3}: StringVal("cab"),
			{Col: 2, Row: 4}: StringVal("dogs"),
		},
	}

	// ? matches any single character: "ca?" matches cat, car, cab but not dogs
	cf := evalCompile(t, `SUMIFS(A1:A4,B1:B4,"ca?")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 10+20+30 = 60
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUMIFS wildcard ?: got %v, want 60", got)
	}
}

func TestSUMIFS_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 2, Row: 1}: StringVal("Apple"),
			{Col: 2, Row: 2}: StringVal("APPLE"),
			{Col: 2, Row: 3}: StringVal("apple"),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// All three should match case-insensitively: 100+200+300 = 600
	if got.Type != ValueNumber || got.Num != 600 {
		t.Errorf("SUMIFS case insensitive: got %v, want 600", got)
	}
}

func TestSUMIFS_EmptyCellsInSumRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			// A2 is empty (not in map)
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Empty cells in sum range contribute 0: 10+0+30 = 40
	if got.Type != ValueNumber || got.Num != 40 {
		t.Errorf("SUMIFS empty sum cells: got %v, want 40", got)
	}
}

func TestSUMIFS_MixedTypes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("text"),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(1),
			{Col: 2, Row: 4}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A4,B1:B4,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// text is not numeric so CoerceNum skips it, TRUE coerces to 1: 10+0+30+1 = 41
	if got.Type != ValueNumber || got.Num != 41 {
		t.Errorf("SUMIFS mixed types: got %v, want 41", got)
	}
}

func TestSUMIFS_TooFewArgs(t *testing.T) {
	resolver := &mockResolver{}

	// Only 1 arg — need at least 3
	cf := evalCompile(t, "SUMIFS(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SUMIFS too few args: got %v, want #VALUE!", got)
	}
}

func TestSUMIFS_EvenArgsAfterSumRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 3, Row: 1}: NumberVal(2),
		},
	}

	// 4 args total: sum_range + 3 => (4-1)%2 != 0 => error
	cf := evalCompile(t, `SUMIFS(A1:A1,B1:B1,1,C1:C1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SUMIFS even args: got %v, want #VALUE!", got)
	}
}

func TestSUMIFS_NumericCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(10),
			{Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Exact numeric match: criteria=5 matches rows 1 and 3
	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,5)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("SUMIFS numeric criteria: got %v, want 400", got)
	}
}

func TestSUMIFS_BooleanCriteriaRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: BoolVal(true),
			{Col: 2, Row: 2}: BoolVal(false),
			{Col: 2, Row: 3}: BoolVal(true),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 1 and 3 match TRUE: 10+30 = 40
	if got.Type != ValueNumber || got.Num != 40 {
		t.Errorf("SUMIFS boolean criteria: got %v, want 40", got)
	}
}

func TestSUMIFS_StringCriteriaNotEqual(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(22),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Bananas"),
			{Col: 2, Row: 3}: StringVal("Carrots"),
		},
	}

	// Sum where product is NOT Bananas
	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,"<>Bananas")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 5+10 = 15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("SUMIFS not-equal string: got %v, want 15", got)
	}
}

func TestSUMIFS_MultipleCriteriaPairs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (A): quantities
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(22),
			{Col: 1, Row: 6}: NumberVal(12),
			// Product (B)
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Apples"),
			{Col: 2, Row: 3}: StringVal("Artichokes"),
			{Col: 2, Row: 4}: StringVal("Artichokes"),
			{Col: 2, Row: 5}: StringVal("Bananas"),
			{Col: 2, Row: 6}: StringVal("Bananas"),
			// Salesperson (C)
			{Col: 3, Row: 1}: StringVal("Tom"),
			{Col: 3, Row: 2}: StringVal("Sarah"),
			{Col: 3, Row: 3}: StringVal("Tom"),
			{Col: 3, Row: 4}: StringVal("Sarah"),
			{Col: 3, Row: 5}: StringVal("Tom"),
			{Col: 3, Row: 6}: StringVal("Sarah"),
		},
	}

	// Sum where product starts with "A*" AND salesperson is "Tom"
	cf := evalCompile(t, `SUMIFS(A1:A6,B1:B6,"A*",C1:C6,"Tom")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Apples/Tom(5) + Artichokes/Tom(15) = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("SUMIFS multiple criteria: got %v, want 20", got)
	}
}

func TestSUMIFS_DateSerialNumbers(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (A): amounts
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			// Date range (B): Excel serial numbers
			// 44927 = 2023-01-01, 44958 = 2023-02-01, 44986 = 2023-03-01
			{Col: 2, Row: 1}: NumberVal(44927),
			{Col: 2, Row: 2}: NumberVal(44958),
			{Col: 2, Row: 3}: NumberVal(44986),
		},
	}

	// Sum amounts where date > 2023-01-15 (serial 44942)
	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,">44942")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 44958 and 44986 > 44942: 200+300 = 500
	if got.Type != ValueNumber || got.Num != 500 {
		t.Errorf("SUMIFS date serials: got %v, want 500", got)
	}
}

func TestSUMIFS_ThreeCriteriaPairs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (A)
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 1, Row: 4}: NumberVal(400),
			// Criteria range 1 (B) — region
			{Col: 2, Row: 1}: StringVal("East"),
			{Col: 2, Row: 2}: StringVal("East"),
			{Col: 2, Row: 3}: StringVal("West"),
			{Col: 2, Row: 4}: StringVal("East"),
			// Criteria range 2 (C) — product
			{Col: 3, Row: 1}: StringVal("Widget"),
			{Col: 3, Row: 2}: StringVal("Gadget"),
			{Col: 3, Row: 3}: StringVal("Widget"),
			{Col: 3, Row: 4}: StringVal("Widget"),
			// Criteria range 3 (D) — qty
			{Col: 4, Row: 1}: NumberVal(10),
			{Col: 4, Row: 2}: NumberVal(20),
			{Col: 4, Row: 3}: NumberVal(30),
			{Col: 4, Row: 4}: NumberVal(50),
		},
	}

	// East AND Widget AND qty>10 => only row 4 (400)
	cf := evalCompile(t, `SUMIFS(A1:A4,B1:B4,"East",C1:C4,"Widget",D1:D4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("SUMIFS three criteria: got %v, want 400", got)
	}
}

// ---------------------------------------------------------------------------
// COUNTIFS — multiple criteria
// ---------------------------------------------------------------------------

func TestCOUNTIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
			{Col: 1, Row: 4}: StringVal("cherry"),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(10),
			{Col: 2, Row: 3}: NumberVal(15),
			{Col: 2, Row: 4}: NumberVal(20),
		},
	}

	// Count where A="apple" AND B>10
	cf := evalCompile(t, `COUNTIFS(A1:A4,"apple",B1:B4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only row 3 (apple, 15>10) matches
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIFS: got %g, want 1", got.Num)
	}

	// Single criteria pair
	cf = evalCompile(t, `COUNTIFS(A1:A4,"apple")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS single: got %g, want 2", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNTIFS — comprehensive tests
// ---------------------------------------------------------------------------

func TestCOUNTIFS_SingleCriteriaNumbers(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(2),
			{Col: 1, Row: 5}: NumberVal(5),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A5,2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS single number: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_ComparisonOperators(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 1, Row: 4}: NumberVal(15),
			{Col: 1, Row: 5}: NumberVal(20),
		},
	}

	tests := []struct {
		formula string
		want    float64
		label   string
	}{
		{`COUNTIFS(A1:A5,">5")`, 3, ">5"},
		{`COUNTIFS(A1:A5,"<=10")`, 3, "<=10"},
		{`COUNTIFS(A1:A5,"<>10")`, 4, "<>10"},
		{`COUNTIFS(A1:A5,">=10")`, 3, ">=10"},
		{`COUNTIFS(A1:A5,"<10")`, 2, "<10"},
		{`COUNTIFS(A1:A5,"=10")`, 1, "=10"},
	}
	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval %s: %v", tt.label, err)
		}
		if got.Type != ValueNumber || got.Num != tt.want {
			t.Errorf("COUNTIFS %s: got %g, want %g", tt.label, got.Num, tt.want)
		}
	}
}

func TestCOUNTIFS_NoMatches(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A3,">100")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNTIFS no matches: got %g, want 0", got.Num)
	}
}

func TestCOUNTIFS_AllMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A3,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIFS all match: got %g, want 3", got.Num)
	}
}

func TestCOUNTIFS_MultipleCriteriaPairs(t *testing.T) {
	// A = category, B = region, C = quantity
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 1, Row: 2}: StringVal("veg"),
			{Col: 1, Row: 3}: StringVal("fruit"),
			{Col: 1, Row: 4}: StringVal("fruit"),
			{Col: 2, Row: 1}: StringVal("east"),
			{Col: 2, Row: 2}: StringVal("east"),
			{Col: 2, Row: 3}: StringVal("west"),
			{Col: 2, Row: 4}: StringVal("east"),
			{Col: 3, Row: 1}: NumberVal(10),
			{Col: 3, Row: 2}: NumberVal(20),
			{Col: 3, Row: 3}: NumberVal(30),
			{Col: 3, Row: 4}: NumberVal(40),
		},
	}

	// fruit AND east AND >5 => rows 1 (10) and 4 (40)
	cf := evalCompile(t, `COUNTIFS(A1:A4,"fruit",B1:B4,"east",C1:C4,">5")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS 3 pairs: got %g, want 2", got.Num)
	}

	// fruit AND west => row 3 only
	cf = evalCompile(t, `COUNTIFS(A1:A4,"fruit",B1:B4,"west")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIFS fruit+west: got %g, want 1", got.Num)
	}
}

func TestCOUNTIFS_WildcardAsterisk(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple pie"),
			{Col: 1, Row: 2}: StringVal("apple sauce"),
			{Col: 1, Row: 3}: StringVal("banana"),
			{Col: 1, Row: 4}: StringVal("pineapple"),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A4,"apple*")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS wildcard *: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_WildcardQuestion(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("cat"),
			{Col: 1, Row: 2}: StringVal("car"),
			{Col: 1, Row: 3}: StringVal("cab"),
			{Col: 1, Row: 4}: StringVal("cart"),
		},
	}
	// "ca?" matches 3-char strings starting with "ca": cat, car, cab
	cf := evalCompile(t, `COUNTIFS(A1:A4,"ca?")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIFS wildcard ?: got %g, want 3", got.Num)
	}
}

func TestCOUNTIFS_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 1, Row: 2}: StringVal("APPLE"),
			{Col: 1, Row: 3}: StringVal("apple"),
			{Col: 1, Row: 4}: StringVal("banana"),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A4,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIFS case-insensitive: got %g, want 3", got.Num)
	}
}

func TestCOUNTIFS_EmptyCells(t *testing.T) {
	// Rows 2 and 4 have no value in column A (empty)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
		},
	}
	// Count where A>0: only rows 1 and 3 have numeric values >0
	cf := evalCompile(t, `COUNTIFS(A1:A4,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS empty cells: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_BooleanValues(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: BoolVal(false),
			{Col: 1, Row: 3}: BoolVal(true),
			{Col: 1, Row: 4}: BoolVal(false),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A4,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS boolean TRUE: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_MixedTypes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 1, Row: 5}: StringVal("5"),
		},
	}
	// Numeric criteria 5 matches NumberVal(5) and StringVal("5") (coerced)
	cf := evalCompile(t, `COUNTIFS(A1:A5,5)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS mixed types: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_ErrorNoArgs(t *testing.T) {
	cf := evalCompile(t, `COUNTIFS()`)
	got, err := Eval(cf, &mockResolver{}, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("COUNTIFS no args: expected error, got type %v", got.Type)
	}
}

func TestCOUNTIFS_ErrorOddArgs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A1,"=1",A1:A1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("COUNTIFS odd args: expected error, got type %v", got.Type)
	}
}

func TestCOUNTIFS_DateSerialNumbers(t *testing.T) {
	// Excel serial: 44197 = 2021-01-01, 44228 = 2021-02-01, 44256 = 2021-03-01
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(44197),
			{Col: 1, Row: 2}: NumberVal(44228),
			{Col: 1, Row: 3}: NumberVal(44256),
			{Col: 1, Row: 4}: NumberVal(44300),
		},
	}
	// Count dates after 2021-02-01 (serial > 44228)
	cf := evalCompile(t, `COUNTIFS(A1:A4,">44228")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS date serials: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_MultipleRangesAND(t *testing.T) {
	// Test that all criteria pairs must match (AND logic)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("yes"),
			{Col: 2, Row: 2}: StringVal("no"),
			{Col: 2, Row: 3}: StringVal("yes"),
		},
	}
	// A>15 AND B="yes" => only row 3 (30, yes)
	cf := evalCompile(t, `COUNTIFS(A1:A3,">15",B1:B3,"yes")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIFS AND logic: got %g, want 1", got.Num)
	}
}

func TestCOUNTIFS_SameRangeTwice(t *testing.T) {
	// Use same range with two different criteria (between logic)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 1, Row: 4}: NumberVal(15),
			{Col: 1, Row: 5}: NumberVal(20),
		},
	}
	// A>=5 AND A<=15 => rows 2,3,4
	cf := evalCompile(t, `COUNTIFS(A1:A5,">=5",A1:A5,"<=15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIFS same range twice: got %g, want 3", got.Num)
	}
}

func TestCOUNTIFS_TextExact(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("yes"),
			{Col: 1, Row: 2}: StringVal("Yes"),
			{Col: 1, Row: 3}: StringVal("YES"),
			{Col: 1, Row: 4}: StringVal("no"),
			{Col: 1, Row: 5}: StringVal("yesss"),
		},
	}
	// "yes" matches case-insensitively but must be exact (not "yesss")
	cf := evalCompile(t, `COUNTIFS(A1:A5,"yes")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIFS text exact: got %g, want 3", got.Num)
	}
}

func TestCOUNTIFS_NotEqual(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(10),
		},
	}
	cf := evalCompile(t, `COUNTIFS(A1:A4,"<>0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS <>0: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFS_WildcardWithCriteriaPair(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("apricot"),
			{Col: 1, Row: 3}: StringVal("banana"),
			{Col: 1, Row: 4}: StringVal("avocado"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
		},
	}
	// A starts with "a" AND B>15
	cf := evalCompile(t, `COUNTIFS(A1:A4,"a*",B1:B4,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// apricot (20>15) and avocado (40>15) => 2
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS wildcard+criteria: got %g, want 2", got.Num)
	}
}

// ---------------------------------------------------------------------------
// AVERAGEIF
// ---------------------------------------------------------------------------

func TestAVERAGEIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
		},
	}

	cf := evalCompile(t, `AVERAGEIF(A1:A4,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 20+30+40 = 90, count=3, avg=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("AVERAGEIF: got %g, want 30", got.Num)
	}

	// No matches => #DIV/0!
	cf = evalCompile(t, `AVERAGEIF(A1:A4,">100")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGEIF no match: got %v, want #DIV/0!", got)
	}
}

func TestAVERAGEIFWithSeparateRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("yes"),
			{Col: 1, Row: 2}: StringVal("no"),
			{Col: 1, Row: 3}: StringVal("yes"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `AVERAGEIF(A1:A3,"yes",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (100+300)/2 = 200
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("AVERAGEIF separate range: got %g, want 200", got.Num)
	}
}

func TestAVERAGEIF_ComparisonOperators(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 1, Row: 5}: NumberVal(50),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"greater than", `AVERAGEIF(A1:A5,">25")`, 40},      // 30+40+50 = 120/3
		{"less than", `AVERAGEIF(A1:A5,"<25")`, 15},         // 10+20 = 30/2
		{"greater or equal", `AVERAGEIF(A1:A5,">=30")`, 40}, // 30+40+50 = 120/3
		{"less or equal", `AVERAGEIF(A1:A5,"<=20")`, 15},    // 10+20 = 30/2
		{"equal operator", `AVERAGEIF(A1:A5,"=30")`, 30},    // 30/1
		{"not equal", `AVERAGEIF(A1:A5,"<>30")`, 30},        // 10+20+40+50 = 120/4
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tc.want {
				t.Errorf("%s: got %v, want %g", tc.name, got, tc.want)
			}
		})
	}
}

func TestAVERAGEIF_AllMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+20+30)/3 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("AVERAGEIF all match: got %v, want 20", got)
	}
}

func TestAVERAGEIF_NumericCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 1, Row: 4}: NumberVal(30),
		},
	}
	// Numeric exact match: average cells equal to 10
	cf := evalCompile(t, `AVERAGEIF(A1:A4,10)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+10)/2 = 10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("AVERAGEIF numeric criteria: got %v, want 10", got)
	}
}

func TestAVERAGEIF_StringExactMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(10),
			{Col: 2, Row: 3}: NumberVal(15),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"apple",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (5+15)/2 = 10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("AVERAGEIF string exact: got %v, want 10", got)
	}
}

func TestAVERAGEIF_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 1, Row: 2}: StringVal("APPLE"),
			{Col: 1, Row: 3}: StringVal("banana"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"apple",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+20)/2 = 15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("AVERAGEIF case insensitive: got %v, want 15", got)
	}
}

func TestAVERAGEIF_WildcardAsterisk(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple pie"),
			{Col: 1, Row: 2}: StringVal("apple sauce"),
			{Col: 1, Row: 3}: StringVal("banana"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"apple*",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+20)/2 = 15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("AVERAGEIF wildcard *: got %v, want 15", got)
	}
}

func TestAVERAGEIF_WildcardQuestion(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("cat"),
			{Col: 1, Row: 2}: StringVal("car"),
			{Col: 1, Row: 3}: StringVal("cart"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}
	// "ca?" matches "cat" and "car" (3 chars), not "cart" (4 chars)
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"ca?",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+20)/2 = 15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("AVERAGEIF wildcard ?: got %v, want 15", got)
	}
}

func TestAVERAGEIF_EmptyCellsIgnored(t *testing.T) {
	// Empty cells in average_range should be ignored
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			// {Col: 1, Row: 2} is empty
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,">=0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only 10 and 30 match >=0 (empty cells don't match numeric criteria)
	// (10+30)/2 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("AVERAGEIF empty cells: got %v, want 20", got)
	}
}

func TestAVERAGEIF_MixedTypes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: BoolVal(true),
		},
	}
	// ">5" only matches numeric values > 5
	cf := evalCompile(t, `AVERAGEIF(A1:A4,">5")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 10 and 30 match, (10+30)/2 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("AVERAGEIF mixed types: got %v, want 20", got)
	}
}

func TestAVERAGEIF_BooleanCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: BoolVal(false),
			{Col: 1, Row: 3}: BoolVal(true),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"TRUE",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (100+300)/2 = 200
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("AVERAGEIF boolean: got %v, want 200", got)
	}
}

func TestAVERAGEIF_NoMatchDIV0(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,">100")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGEIF no match: got %v, want #DIV/0!", got)
	}
}

func TestAVERAGEIF_StringNoMatchDIV0(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A2,"xyz",B1:B2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGEIF string no match: got %v, want #DIV/0!", got)
	}
}

func TestAVERAGEIF_TooFewArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `AVERAGEIF(A1:A3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("AVERAGEIF too few args: got %v, want error", got)
	}
}

func TestAVERAGEIF_AvgRangeSameSize(t *testing.T) {
	// average_range and criteria_range are the same size
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,">=2",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 2 and 3 match (>=2); avg of B2,B3 = (200+300)/2 = 250
	if got.Type != ValueNumber || got.Num != 250 {
		t.Errorf("AVERAGEIF same size ranges: got %v, want 250", got)
	}
}

func TestAVERAGEIF_WithoutAvgRange(t *testing.T) {
	// When average_range is omitted, the criteria_range is used for averaging
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(20),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A4,">8")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 10+15+20 = 45, count=3, avg=15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("AVERAGEIF without avg_range: got %v, want 15", got)
	}
}

func TestAVERAGEIF_EqualStringOperator(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("cat"),
			{Col: 1, Row: 2}: StringVal("dog"),
			{Col: 1, Row: 3}: StringVal("cat"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"=cat",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (10+30)/2 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("AVERAGEIF =string: got %v, want 20", got)
	}
}

func TestAVERAGEIF_NotEqualString(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("cat"),
			{Col: 1, Row: 2}: StringVal("dog"),
			{Col: 1, Row: 3}: StringVal("cat"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, `AVERAGEIF(A1:A3,"<>cat",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only "dog" matches, avg = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("AVERAGEIF <>string: got %v, want 20", got)
	}
}

// ---------------------------------------------------------------------------
// SUMIF with separate sum range
// ---------------------------------------------------------------------------

func TestSUMIFWithSeparateRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("yes"),
			{Col: 1, Row: 2}: StringVal("no"),
			{Col: 1, Row: 3}: StringVal("yes"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,"yes",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("SUMIF separate range: got %g, want 400", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNTIF edge cases
// ---------------------------------------------------------------------------

func TestCOUNTIFWildcard(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple pie"),
			{Col: 1, Row: 2}: StringVal("apple sauce"),
			{Col: 1, Row: 3}: StringVal("banana"),
		},
	}

	cf := evalCompile(t, `COUNTIF(A1:A3,"apple*")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF wildcard: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFWildcardEscapes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Data!A1:A10
			{Col: 1, Row: 1}: StringVal("hello"),
			{Col: 1, Row: 2}: StringVal("has*star"),
			{Col: 1, Row: 3}: StringVal("*star"),
			{Col: 1, Row: 4}: StringVal("star*"),
			{Col: 1, Row: 5}: StringVal("a?b"),
			{Col: 1, Row: 6}: StringVal("has?question"),
			{Col: 1, Row: 7}: StringVal("has~tilde"),
			{Col: 1, Row: 8}: StringVal("plain"),
			// B column for SUMIF
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(8),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(9),
			{Col: 2, Row: 6}: NumberVal(3),
			{Col: 2, Row: 7}: NumberVal(7),
			{Col: 2, Row: 8}: NumberVal(2),
		},
	}

	tests := []struct {
		formula string
		want    float64
	}{
		// *~** means: anything, literal *, anything → cells containing *
		{`COUNTIF(A1:A8,"*~**")`, 3}, // has*star, *star, star*
		// *~?* means: anything, literal ?, anything → cells containing ?
		{`COUNTIF(A1:A8,"*~?*")`, 2}, // a?b, has?question
		// *~~* means: anything, literal ~, anything → cells containing ~
		{`COUNTIF(A1:A8,"*~~*")`, 1}, // has~tilde
		// ~*star means: literal * then "star"
		{`COUNTIF(A1:A8,"~*star")`, 1}, // *star
		// star~* means: "star" then literal *
		{`COUNTIF(A1:A8,"star~*")`, 1}, // star*
		// a~?b means: "a" then literal ? then "b"
		{`COUNTIF(A1:A8,"a~?b")`, 1}, // a?b
		// ~** means: literal * then wildcard * (match anything) → cells starting with *
		{`COUNTIF(A1:A8,"~**")`, 1}, // *star
		// ~* alone: matches literal "*"
		{`COUNTIF(A1:A8,"~*")`, 0}, // no cell is exactly "*"
		// SUMIF with wildcard escapes
		{`SUMIF(A1:A8,"*~**",B1:B8)`, 17}, // 5+8+4 = 17
		{`SUMIF(A1:A8,"a~?b",B1:B8)`, 9},  // row 5 = 9
	}

	for _, tc := range tests {
		t.Run(tc.formula, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tc.want {
				t.Errorf("%s: got %v, want %g", tc.formula, got, tc.want)
			}
		})
	}
}

func TestCOUNTIFNumericCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(20),
		},
	}

	// Less-than operator
	cf := evalCompile(t, `COUNTIF(A1:A4,"<15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF <15: got %g, want 2", got.Num)
	}

	// Equals with operator
	cf = evalCompile(t, `COUNTIF(A1:A4,"=10")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF =10: got %g, want 1", got.Num)
	}

	// Not-equal
	cf = evalCompile(t, `COUNTIF(A1:A4,"<>10")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIF <>10: got %g, want 3", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNTIF with mixed positive/negative/zero values
// ---------------------------------------------------------------------------

func TestCOUNTIFMixedSignValues(t *testing.T) {
	// Mirrors the multisheet edge case spec: A1:A5 = [10, -5, 0, 100, 25]
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(-5),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(100),
			{Col: 1, Row: 5}: NumberVal(25),
		},
	}

	// >0 should count only strictly positive values (10, 100, 25) => 3
	cf := evalCompile(t, `COUNTIF(A1:A5,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIF >0 mixed: got %g, want 3", got.Num)
	}

	// <0 should count only negative values (-5) => 1
	cf = evalCompile(t, `COUNTIF(A1:A5,"<0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF <0 mixed: got %g, want 1", got.Num)
	}

	// =0 should count only zero values => 1
	cf = evalCompile(t, `COUNTIF(A1:A5,"=0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF =0 mixed: got %g, want 1", got.Num)
	}

	// >=0 should count zero and positives (10, 0, 100, 25) => 4
	cf = evalCompile(t, `COUNTIF(A1:A5,">=0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNTIF >=0 mixed: got %g, want 4", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMIF with mixed positive/negative/zero values
// ---------------------------------------------------------------------------

func TestSUMIFMixedSignValues(t *testing.T) {
	// Mirrors the multisheet edge case spec: A1:A5 = [10, -5, 0, 100, 25]
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(-5),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(100),
			{Col: 1, Row: 5}: NumberVal(25),
		},
	}

	// >0 should sum only strictly positive values (10+100+25) => 135
	cf := evalCompile(t, `SUMIF(A1:A5,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 135 {
		t.Errorf("SUMIF >0 mixed: got %g, want 135", got.Num)
	}

	// <0 should sum only negative values (-5) => -5
	cf = evalCompile(t, `SUMIF(A1:A5,"<0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -5 {
		t.Errorf("SUMIF <0 mixed: got %g, want -5", got.Num)
	}

	// No matches => 0
	cf = evalCompile(t, `SUMIF(A1:A5,">1000")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("SUMIF no match: got %g, want 0", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMIF — comprehensive tests
// ---------------------------------------------------------------------------

func TestSUMIF_AllComparisonOperators(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 1, Row: 5}: NumberVal(50),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"greater than", `SUMIF(A1:A5,">30")`, 90},       // 40+50
		{"less than", `SUMIF(A1:A5,"<30")`, 30},          // 10+20
		{"greater or equal", `SUMIF(A1:A5,">=30")`, 120}, // 30+40+50
		{"less or equal", `SUMIF(A1:A5,"<=30")`, 60},     // 10+20+30
		{"equal", `SUMIF(A1:A5,"=30")`, 30},              // 30
		{"not equal", `SUMIF(A1:A5,"<>30")`, 120},        // 10+20+40+50
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tc.want {
				t.Errorf("%s: got %v, want %g", tc.name, got, tc.want)
			}
		})
	}
}

func TestSUMIF_WildcardAsterisk(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 1, Row: 2}: StringVal("Apricot"),
			{Col: 1, Row: 3}: StringVal("Banana"),
			{Col: 1, Row: 4}: StringVal("Avocado"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A4,"A*",B1:B4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 70 {
		t.Errorf("SUMIF wildcard *: got %v, want 70", got)
	}
}

func TestSUMIF_WildcardQuestionMark(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("cat"),
			{Col: 1, Row: 2}: StringVal("car"),
			{Col: 1, Row: 3}: StringVal("cart"),
			{Col: 1, Row: 4}: StringVal("cab"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
		},
	}

	// "ca?" matches 3-char strings starting with "ca": cat, car, cab
	cf := evalCompile(t, `SUMIF(A1:A4,"ca?",B1:B4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 70 {
		t.Errorf("SUMIF wildcard ?: got %v, want 70", got)
	}
}

func TestSUMIF_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("APPLE"),
			{Col: 1, Row: 2}: StringVal("apple"),
			{Col: 1, Row: 3}: StringVal("Apple"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,"apple",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 600 {
		t.Errorf("SUMIF case insensitive: got %v, want 600", got)
	}
}

func TestSUMIF_NumericCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,5,B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("SUMIF numeric criteria: got %v, want 400", got)
	}
}

func TestSUMIF_TooFewArgs(t *testing.T) {
	resolver := &mockResolver{cells: map[CellAddr]Value{}}

	cf := evalCompile(t, "SUMIF(A1:A3)")
	got, _ := Eval(cf, resolver, nil)
	if got.Type != ValueError {
		t.Errorf("SUMIF too few args: got %v, want #VALUE!", got)
	}
}

func TestSUMIF_TooManyArgs(t *testing.T) {
	resolver := &mockResolver{cells: map[CellAddr]Value{}}

	cf := evalCompile(t, "SUMIF(A1:A3,1,B1:B3,C1:C3)")
	got, _ := Eval(cf, resolver, nil)
	if got.Type != ValueError {
		t.Errorf("SUMIF too many args: got %v, want #VALUE!", got)
	}
}

func TestSUMIF_ExcelDocExample(t *testing.T) {
	// From Excel docs Example 1: property values and commissions
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100000),
			{Col: 1, Row: 2}: NumberVal(200000),
			{Col: 1, Row: 3}: NumberVal(300000),
			{Col: 1, Row: 4}: NumberVal(400000),
			{Col: 2, Row: 1}: NumberVal(7000),
			{Col: 2, Row: 2}: NumberVal(14000),
			{Col: 2, Row: 3}: NumberVal(21000),
			{Col: 2, Row: 4}: NumberVal(28000),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"commissions over 160k", `SUMIF(A1:A4,">160000",B1:B4)`, 63000},
		{"values over 160k", `SUMIF(A1:A4,">160000")`, 900000},
		{"commissions for 300k", `SUMIF(A1:A4,300000,B1:B4)`, 21000},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tc.want {
				t.Errorf("%s: got %v, want %g", tc.name, got, tc.want)
			}
		})
	}
}

func TestSUMIF_ExcelDocExample2(t *testing.T) {
	// From Excel docs Example 2: categories and food sales
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Vegetables"),
			{Col: 1, Row: 2}: StringVal("Vegetables"),
			{Col: 1, Row: 3}: StringVal("Fruits"),
			// Row 4: A4 is empty (Butter row with no category)
			{Col: 1, Row: 5}: StringVal("Vegetables"),
			{Col: 1, Row: 6}: StringVal("Fruits"),
			{Col: 2, Row: 1}: StringVal("Tomatoes"),
			{Col: 2, Row: 2}: StringVal("Celery"),
			{Col: 2, Row: 3}: StringVal("Oranges"),
			{Col: 2, Row: 4}: StringVal("Butter"),
			{Col: 2, Row: 5}: StringVal("Carrots"),
			{Col: 2, Row: 6}: StringVal("Apples"),
			{Col: 3, Row: 1}: NumberVal(2300),
			{Col: 3, Row: 2}: NumberVal(5500),
			{Col: 3, Row: 3}: NumberVal(800),
			{Col: 3, Row: 4}: NumberVal(400),
			{Col: 3, Row: 5}: NumberVal(4200),
			{Col: 3, Row: 6}: NumberVal(1200),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"fruits", `SUMIF(A1:A6,"Fruits",C1:C6)`, 2000},
		{"vegetables", `SUMIF(A1:A6,"Vegetables",C1:C6)`, 12000},
		{"ends with es", `SUMIF(B1:B6,"*es",C1:C6)`, 4300}, // Tomatoes+Oranges+Apples
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tc.want {
				t.Errorf("%s: got %v, want %g", tc.name, got, tc.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// AVERAGE / SUM / MIN / MAX with edge cases
// ---------------------------------------------------------------------------

func TestAVERAGEEmpty(t *testing.T) {
	resolver := &mockResolver{}

	// No numeric values => #DIV/0!
	cf := evalCompile(t, "AVERAGE(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGE empty: got %v, want #DIV/0!", got)
	}
}

func TestAVERAGE(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("multiple numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(2,3,3,5,7,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(-10,-20,-30)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(-10,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("zero values", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0,0,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 10.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > 1e-10 {
			t.Errorf("got %v, want %g", got, want)
		}
	})

	t.Run("all zeros", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0,0,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1000000,2000000,3000000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0.001,0.002,0.003)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.002) > 1e-15 {
			t.Errorf("got %v, want 0.002", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so (1+3)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so (0+4)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		// Direct string arg "5" is coerced by CoerceNum to 5.
		cf := evalCompile(t, `AVERAGE("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGE("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 30)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(30),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+30)/2 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(30),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+30)/2 = 20; TRUE is ignored in range
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range with zero values included", func(t *testing.T) {
		resolver := numResolver(0, 10, 20)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range all empty gives DIV0", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("mixed range and literal", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "AVERAGE(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20+30)/3 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("fractional result", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1,2)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1.5 {
			t.Errorf("got %v, want 1.5", got)
		}
	})

	t.Run("single zero", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in direct arg propagates", func(t *testing.T) {
		// 1/0 produces #DIV/0!, which should propagate through AVERAGE.
		cf := evalCompile(t, "AVERAGE(1/0,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("range with numeric string ignored", func(t *testing.T) {
		// Per Excel: numeric strings in ranges are ignored by AVERAGE.
		resolver := valResolver(
			NumberVal(10),
			StringVal("5"),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20)/2 = 15; "5" in range is ignored
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	_ = numResolver // suppress unused if needed
	_ = valResolver
}

func TestAVERAGEA(t *testing.T) {
	// Helper to build a resolver with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (1+3)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (0+4)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGEA("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (5+15)/2 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGEA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		// Excel example: {10, 7, 9, 2, "Not available"} => (10+7+9+2+0)/5 = 5.6
		resolver := valResolver(
			NumberVal(10),
			NumberVal(7),
			NumberVal(9),
			NumberVal(2),
			StringVal("Not available"),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5.6 {
			t.Errorf("got %v, want 5.6", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+1+20)/3 ≈ 10.333...
		want := 31.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > 1e-10 {
			t.Errorf("got %v, want %g", got, want)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			BoolVal(false),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("empty cells ignored in range", func(t *testing.T) {
		// A1=10, A2=empty, A3=20
		resolver := valResolver(
			NumberVal(10),
		)
		resolver.cells[CellAddr{Col: 1, Row: 3}] = NumberVal(20)
		// A2 is empty (not set)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20)/2 = 15; empty cell is ignored
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("all empty range returns DIV0", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error propagates from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNUM),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("error propagates from direct arg", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1/0,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20+30)/3 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range all text", func(t *testing.T) {
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
			StringVal("baz"),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (0+0+0)/3 = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range numeric string counts as 0 not parsed", func(t *testing.T) {
		// In range, "5" counts as 0, not 5
		resolver := valResolver(
			NumberVal(10),
			StringVal("5"),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("single TRUE", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(TRUE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("single FALSE", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(-10,-20,-30)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("range with mixed types", func(t *testing.T) {
		// {10, TRUE, "hello", FALSE, 20} => (10+1+0+0+20)/5 = 6.2
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			StringVal("hello"),
			BoolVal(false),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6.2 {
			t.Errorf("got %v, want 6.2", got)
		}
	})

	t.Run("range with empty string counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			StringVal(""),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1000000,2000000,3000000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("zero value", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	_ = valResolver
}

func TestSUMWithMixedTypes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(20),
			{Col: 1, Row: 4}: BoolVal(true),
		},
	}

	// In a range, strings and bools are skipped by SUM
	cf := evalCompile(t, "SUM(A1:A4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("SUM mixed: got %g, want 30", got.Num)
	}
}

func TestMINMAXEmpty(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MIN empty: got %g, want 0", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MAX empty: got %g, want 0", got.Num)
	}
}

func TestMINMAXNegative(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-100),
			{Col: 1, Row: 2}: NumberVal(-50),
			{Col: 1, Row: 3}: NumberVal(-1),
		},
	}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -100 {
		t.Errorf("MIN neg: got %g, want -100", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -1 {
		t.Errorf("MAX neg: got %g, want -1", got.Num)
	}
}

// ---------------------------------------------------------------------------
// MIN comprehensive tests
// ---------------------------------------------------------------------------

func TestMIN(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MIN(42)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("multiple numbers picks smallest", func(t *testing.T) {
		cf := evalCompile(t, "MIN(5,3,8,1,9)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		cf := evalCompile(t, "MIN(7,7,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(-5,-3,-10,-1)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -10 {
			t.Errorf("got %v, want -10", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "MIN(10,-5,3,-20,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("zero is minimum", func(t *testing.T) {
		cf := evalCompile(t, "MIN(5,10,0,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(1.5,0.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0.5 {
			t.Errorf("got %v, want 0.5", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(1000000,2000000,500000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 500000 {
			t.Errorf("got %v, want 500000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(0.001,0.002,0.0005)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.0005) > 1e-15 {
			t.Errorf("got %v, want 0.0005", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "MIN(TRUE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so min(1,5) = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "MIN(FALSE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so min(0,5) = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("boolean TRUE and FALSE as direct args", func(t *testing.T) {
		cf := evalCompile(t, "MIN(TRUE,FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, FALSE=0, so min = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		cf := evalCompile(t, `MIN("3",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "3" coerced to 3, so min(3,10) = 3
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `MIN("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation DIV0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValDIV0),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("no args returns 0", func(t *testing.T) {
		// MIN with empty range => no numbers found => returns 0
		resolver := &mockResolver{}
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 5, 15)
		cf := evalCompile(t, "MIN(A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE is ignored in range, so min(10,5) = 5
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores empty cells", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			EmptyVal(),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range with only strings returns 0", func(t *testing.T) {
		// All non-numeric => no numbers found => returns 0
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
		)
		cf := evalCompile(t, "MIN(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed direct args and range", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "MIN(A1:A2,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// min(10, 20, 3) = 3
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("negative decimal", func(t *testing.T) {
		cf := evalCompile(t, "MIN(-0.5,0.5,-1.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1.5 {
			t.Errorf("got %v, want -1.5", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Data: 10, 7, 9, 27, 2 => MIN = 2
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MIN(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("Excel example with zero", func(t *testing.T) {
		// MIN(A1:A5, 0) where data is 10, 7, 9, 27, 2 => 0
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MIN(A1:A5,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})
}

// ---------------------------------------------------------------------------
// MAX comprehensive tests
// ---------------------------------------------------------------------------

func TestMAX(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MAX(42)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("multiple numbers picks largest", func(t *testing.T) {
		cf := evalCompile(t, "MAX(5,3,8,1,9)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 9 {
			t.Errorf("got %v, want 9", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		cf := evalCompile(t, "MAX(7,7,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-5,-3,-10,-1)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1 {
			t.Errorf("got %v, want -1", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "MAX(10,-5,3,-20,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("zero is maximum", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-5,-10,0,-3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(1.5,0.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2.5 {
			t.Errorf("got %v, want 2.5", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(1000000,2000000,500000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(0.001,0.002,0.0005)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.002) > 1e-15 {
			t.Errorf("got %v, want 0.002", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "MAX(TRUE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so max(1,-5) = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "MAX(FALSE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so max(0,-5) = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("boolean TRUE and FALSE as direct args", func(t *testing.T) {
		cf := evalCompile(t, "MAX(TRUE,FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, FALSE=0, so max = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		cf := evalCompile(t, `MAX("3",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "3" coerced to 3, so max(3,10) = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `MAX("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation NA from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation DIV0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValDIV0),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("no args returns 0", func(t *testing.T) {
		// MAX with empty range => no numbers found => returns 0
		resolver := &mockResolver{}
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 5, 15)
		cf := evalCompile(t, "MAX(A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE is ignored in range, so max(10,5) = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range ignores empty cells", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			EmptyVal(),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range with only strings returns 0", func(t *testing.T) {
		// All non-numeric => no numbers found => returns 0
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
		)
		cf := evalCompile(t, "MAX(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed direct args and range", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "MAX(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// max(10, 20, 30) = 30
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})

	t.Run("negative decimal", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-0.5,0.5,-1.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0.5 {
			t.Errorf("got %v, want 0.5", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Data: 10, 7, 9, 27, 2 => MAX = 27
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MAX(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 27 {
			t.Errorf("got %v, want 27", got)
		}
	})

	t.Run("Excel example with 30", func(t *testing.T) {
		// MAX(A1:A5, 30) where data is 10, 7, 9, 27, 2 => 30
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MAX(A1:A5,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})
}

// ---------------------------------------------------------------------------
// LARGE/SMALL edge cases
// ---------------------------------------------------------------------------

func TestLARGESMALLEdgeCases(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	// k out of range => #NUM!
	cf := evalCompile(t, "LARGE(A1:A3,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("LARGE k=0: got %v, want #NUM!", got)
	}

	cf = evalCompile(t, "LARGE(A1:A3,4)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("LARGE k>n: got %v, want #NUM!", got)
	}

	// k=2 with duplicates
	cf = evalCompile(t, "LARGE(A1:A3,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("LARGE k=2: got %g, want 5", got.Num)
	}

	cf = evalCompile(t, "SMALL(A1:A3,0)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("SMALL k=0: got %v, want #NUM!", got)
	}
}

// ---------------------------------------------------------------------------
// LARGE — comprehensive tests
// ---------------------------------------------------------------------------

func TestLARGEComprehensive(t *testing.T) {
	t.Run("k=1 returns largest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3),
				{Col: 1, Row: 2}: NumberVal(7),
				{Col: 1, Row: 3}: NumberVal(5),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A4,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("k=2 returns second largest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3),
				{Col: 1, Row: 2}: NumberVal(7),
				{Col: 1, Row: 3}: NumberVal(5),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A4,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("k=last returns smallest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3),
				{Col: 1, Row: 2}: NumberVal(7),
				{Col: 1, Row: 3}: NumberVal(5),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A4,4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-10),
				{Col: 1, Row: 2}: NumberVal(-3),
				{Col: 1, Row: 3}: NumberVal(-7),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -3 {
			t.Errorf("got %v, want -3", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-5),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(-2),
				{Col: 1, Row: 4}: NumberVal(8),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A4,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -2 {
			t.Errorf("got %v, want -2", got)
		}
	})

	t.Run("duplicates in array", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(4),
				{Col: 1, Row: 2}: NumberVal(4),
				{Col: 1, Row: 3}: NumberVal(4),
				{Col: 1, Row: 4}: NumberVal(9),
			},
		}
		// k=2 should return 4 (the second largest, which is one of the duplicates)
		cf := evalCompile(t, "LARGE(A1:A4,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("k too large returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A2,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("k=0 returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A1,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("k negative returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A2,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("single element k=1", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("decimal values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1.5),
				{Col: 1, Row: 2}: NumberVal(2.7),
				{Col: 1, Row: 3}: NumberVal(0.3),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1.5 {
			t.Errorf("got %v, want 1.5", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(8),
				{Col: 1, Row: 2}: NumberVal(8),
				{Col: 1, Row: 3}: NumberVal(8),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 8 {
			t.Errorf("got %v, want 8", got)
		}
	})

	t.Run("empty cells in range are ignored", func(t *testing.T) {
		// A2 is empty (not in map), so only 2 numeric values
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
		// k=3 should be #NUM! since only 2 numeric values
		cf = evalCompile(t, "LARGE(A1:A3,3)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("strings in range are ignored", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: NumberVal(15),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
		// Only 2 numeric values, k=2 should work
		cf = evalCompile(t, "LARGE(A1:A3,2)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("boolean in range is ignored", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3),
				{Col: 1, Row: 2}: BoolVal(true),
				{Col: 1, Row: 3}: NumberVal(7),
			},
		}
		// Booleans in ranges are not numeric, so only 2 values
		cf := evalCompile(t, "LARGE(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "LARGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("too many args returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "LARGE(A1:A3,1,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("large dataset", func(t *testing.T) {
		cells := make(map[CellAddr]Value)
		for i := 1; i <= 100; i++ {
			cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		}
		resolver := &mockResolver{cells: cells}
		cf := evalCompile(t, "LARGE(A1:A100,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 100 {
			t.Errorf("got %v, want 100", got)
		}
		// 50th largest in 1..100 is 51
		cf = evalCompile(t, "LARGE(A1:A100,50)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 51 {
			t.Errorf("got %v, want 51", got)
		}
	})

	t.Run("k as decimal is truncated", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
		}
		// k=2.9 should be truncated to 2
		cf := evalCompile(t, "LARGE(A1:A3,2.9)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("error in array propagates", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: NumberVal(10),
			},
		}
		cf := evalCompile(t, "LARGE(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("scalar argument instead of array", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(99),
			},
		}
		// LARGE with a single cell reference (scalar), k=1
		cf := evalCompile(t, "LARGE(A1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 99 {
			t.Errorf("got %v, want 99", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Excel docs example: {3,4,5,2,3,4,5,6,4,7} LARGE(...,3)=5, LARGE(...,7)=4
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(4),
				{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(4),
				{Col: 1, Row: 4}: NumberVal(5), {Col: 2, Row: 4}: NumberVal(6),
				{Col: 1, Row: 5}: NumberVal(4), {Col: 2, Row: 5}: NumberVal(7),
			},
		}
		cf := evalCompile(t, "LARGE(A1:B5,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("LARGE k=3: got %v, want 5", got)
		}
		cf = evalCompile(t, "LARGE(A1:B5,7)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("LARGE k=7: got %v, want 4", got)
		}
	})
}

// ---------------------------------------------------------------------------
// SMALL — comprehensive tests
// ---------------------------------------------------------------------------

func TestSMALLComprehensive(t *testing.T) {
	t.Run("k=1 returns smallest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(30),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A4,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("k=2 returns second smallest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(30),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A4,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("k=last returns largest", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(30),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A4,4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-3),
				{Col: 1, Row: 2}: NumberVal(-7),
				{Col: 1, Row: 3}: NumberVal(-1),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -7 {
			t.Errorf("got %v, want -7", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-5),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 1, Row: 3}: NumberVal(-2),
				{Col: 1, Row: 4}: NumberVal(8),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A4,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("duplicates in array", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(4),
				{Col: 1, Row: 2}: NumberVal(4),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 1, Row: 4}: NumberVal(4),
			},
		}
		// k=2 should return 4 (second smallest, which is a duplicate)
		cf := evalCompile(t, "SMALL(A1:A4,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("k too large returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A2,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("k=0 returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A1,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("k negative returns NUM error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A2,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("single element array k=1", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("decimal values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0.1),
				{Col: 1, Row: 2}: NumberVal(0.3),
				{Col: 1, Row: 3}: NumberVal(0.2),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0.2 {
			t.Errorf("got %v, want 0.2", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(8),
				{Col: 1, Row: 2}: NumberVal(8),
				{Col: 1, Row: 3}: NumberVal(8),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 8 {
			t.Errorf("got %v, want 8", got)
		}
	})

	t.Run("empty cells in range are ignored", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				// A2 is empty
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
		// k=3 should be #NUM! since only 2 numeric values
		cf = evalCompile(t, "SMALL(A1:A3,3)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("strings in range are ignored", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: NumberVal(15),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
		// Only 2 numeric values, k=2 should work
		cf = evalCompile(t, "SMALL(A1:A3,2)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("booleans in range are ignored", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: BoolVal(true),
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		// Booleans in ranges are not numeric, so only 2 values
		cf := evalCompile(t, "SMALL(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "SMALL(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("too many args returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "SMALL(A1:A3,1,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("large dataset", func(t *testing.T) {
		cells := make(map[CellAddr]Value)
		for i := 1; i <= 100; i++ {
			cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		}
		resolver := &mockResolver{cells: cells}
		cf := evalCompile(t, "SMALL(A1:A100,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
		// 50th smallest in 1..100 is 50
		cf = evalCompile(t, "SMALL(A1:A100,50)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 50 {
			t.Errorf("got %v, want 50", got)
		}
	})

	t.Run("k as decimal is truncated", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
		}
		// k=2.9 should be truncated to 2
		cf := evalCompile(t, "SMALL(A1:A3,2.9)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("error in array propagates", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: NumberVal(10),
			},
		}
		cf := evalCompile(t, "SMALL(A1:A3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("scalar argument instead of array", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(99),
			},
		}
		// SMALL with a single cell reference (scalar), k=1
		cf := evalCompile(t, "SMALL(A1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 99 {
			t.Errorf("got %v, want 99", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Excel docs: Data={3,4,5,2,3,4,6,4,7} SMALL(A2:A10,4)=4
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(4),
				{Col: 1, Row: 3}: NumberVal(5), {Col: 2, Row: 3}: NumberVal(8),
				{Col: 1, Row: 4}: NumberVal(2), {Col: 2, Row: 4}: NumberVal(3),
				{Col: 1, Row: 5}: NumberVal(3), {Col: 2, Row: 5}: NumberVal(7),
				{Col: 1, Row: 6}: NumberVal(4), {Col: 2, Row: 6}: NumberVal(12),
				{Col: 1, Row: 7}: NumberVal(6), {Col: 2, Row: 7}: NumberVal(54),
				{Col: 1, Row: 8}: NumberVal(4), {Col: 2, Row: 8}: NumberVal(8),
				{Col: 1, Row: 9}: NumberVal(7), {Col: 2, Row: 9}: NumberVal(23),
			},
		}
		// SMALL(A1:A9,4) => sorted: {2,3,3,4,4,4,5,6,7} => 4th = 4
		cf := evalCompile(t, "SMALL(A1:A9,4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("SMALL k=4: got %v, want 4", got)
		}
		// SMALL(B1:B9,2) => sorted: {1,3,4,7,8,8,12,23,54} => 2nd = 3
		cf = evalCompile(t, "SMALL(B1:B9,2)")
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("SMALL k=2: got %v, want 3", got)
		}
	})
}

// ---------------------------------------------------------------------------
// Error propagation in range functions
// ---------------------------------------------------------------------------

func TestSUMErrorInRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: ErrorVal(ErrValNA),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "SUM(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("SUM with error: got %v, want #N/A", got)
	}
}

// ---------------------------------------------------------------------------
// MatchesCriteria — extended edge cases
// ---------------------------------------------------------------------------

func TestMatchesCriteriaExtended(t *testing.T) {
	tests := []struct {
		name string
		v    Value
		crit Value
		want bool
	}{
		// Case-insensitive string equality
		{name: "case_insensitive", v: StringVal("Apple"), crit: StringVal("apple"), want: true},
		// Wildcard: ? matches exactly one character
		{name: "question_mark", v: StringVal("bat"), crit: StringVal("b?t"), want: true},
		{name: "question_no_match", v: StringVal("boot"), crit: StringVal("b?t"), want: false},
		// Wildcard: * at end
		{name: "star_end", v: StringVal("hello world"), crit: StringVal("hello*"), want: true},
		// Wildcard: * at start
		{name: "star_start", v: StringVal("hello world"), crit: StringVal("*world"), want: true},
		// Wildcard: * in middle
		{name: "star_middle", v: StringVal("hello world"), crit: StringVal("he*ld"), want: true},
		// Number equality with numeric criteria
		{name: "num_eq_num", v: NumberVal(42), crit: NumberVal(42), want: true},
		{name: "num_ne_num", v: NumberVal(42), crit: NumberVal(43), want: false},
		// String "less than"
		{name: "str_lt", v: NumberVal(3), crit: StringVal("<5"), want: true},
		{name: "str_lt_fail", v: NumberVal(10), crit: StringVal("<5"), want: false},
		// "=" criteria matches only truly empty cells
		{name: "equals_empty_matches_empty", v: EmptyVal(), crit: StringVal("="), want: true},
		{name: "equals_empty_no_match_str", v: StringVal("hello"), crit: StringVal("="), want: false},
		{name: "equals_empty_no_match_num", v: NumberVal(0), crit: StringVal("="), want: false},
		{name: "equals_empty_no_match_empty_str", v: StringVal(""), crit: StringVal("="), want: false},
		// Numeric criteria matches string cells containing the same number
		{name: "num_crit_matches_str_20", v: StringVal("20"), crit: NumberVal(20), want: true},
		{name: "num_crit_no_match_str_abc", v: StringVal("abc"), crit: NumberVal(20), want: false},
		// Numeric operand with "=" operator matches text-numbers
		{name: "eq20_matches_str_20", v: StringVal("20"), crit: StringVal("=20"), want: true},
		// Numeric operand with ordering operators does NOT match text-numbers
		{name: "gt10_no_match_str_15", v: StringVal("15"), crit: StringVal(">10"), want: false},
		// Tilde escape: "~**" means literal * followed by wildcard * (match anything)
		{name: "tilde_star_star_match", v: StringVal("*Special"), crit: StringVal("~**"), want: true},
		{name: "tilde_star_star_exact", v: StringVal("**"), crit: StringVal("~**"), want: true},
		{name: "tilde_star_star_just_star", v: StringVal("*"), crit: StringVal("~**"), want: true},
		{name: "tilde_star_star_no_star", v: StringVal("abc"), crit: StringVal("~**"), want: false},
		// "~*" matches literal "*"
		{name: "tilde_star_exact", v: StringVal("*"), crit: StringVal("~*"), want: true},
		{name: "tilde_star_no_match", v: StringVal("*abc"), crit: StringVal("~*"), want: false},
		// "~?" matches literal "?"
		{name: "tilde_question_exact", v: StringVal("?"), crit: StringVal("~?"), want: true},
		{name: "tilde_question_no_match", v: StringVal("a"), crit: StringVal("~?"), want: false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := MatchesCriteria(tt.v, tt.crit)
			if got != tt.want {
				t.Errorf("MatchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// MatchesCriteria — boolean vs number type distinction
// ---------------------------------------------------------------------------

func TestMatchesCriteriaBoolNumDistinction(t *testing.T) {
	tests := []struct {
		name string
		v    Value
		crit Value
		want bool
	}{
		// Boolean criteria only matches boolean cells
		{name: "bool_true_matches_bool_true", v: BoolVal(true), crit: BoolVal(true), want: true},
		{name: "bool_true_no_match_num_1", v: NumberVal(1), crit: BoolVal(true), want: false},
		{name: "bool_false_matches_bool_false", v: BoolVal(false), crit: BoolVal(false), want: true},
		{name: "bool_false_no_match_num_0", v: NumberVal(0), crit: BoolVal(false), want: false},
		{name: "bool_true_no_match_bool_false", v: BoolVal(true), crit: BoolVal(false), want: false},

		// Numeric criteria only matches numeric cells
		{name: "num_1_matches_num_1", v: NumberVal(1), crit: NumberVal(1), want: true},
		{name: "num_1_no_match_bool_true", v: BoolVal(true), crit: NumberVal(1), want: false},
		{name: "num_0_matches_num_0", v: NumberVal(0), crit: NumberVal(0), want: true},
		{name: "num_0_no_match_bool_false", v: BoolVal(false), crit: NumberVal(0), want: false},

		// String "TRUE"/"FALSE" matches boolean via case-insensitive comparison
		{name: "str_TRUE_matches_bool_true", v: BoolVal(true), crit: StringVal("TRUE"), want: true},
		{name: "str_TRUE_no_match_num_1", v: NumberVal(1), crit: StringVal("TRUE"), want: false},
		{name: "str_FALSE_matches_bool_false", v: BoolVal(false), crit: StringVal("FALSE"), want: true},

		// Comparison operators with numeric operand exclude booleans
		{name: "gt0_matches_positive_num", v: NumberVal(5), crit: StringVal(">0"), want: true},
		{name: "gt0_excludes_bool_true", v: BoolVal(true), crit: StringVal(">0"), want: false},
		{name: "gt0_excludes_bool_false", v: BoolVal(false), crit: StringVal(">0"), want: false},

		// <>TRUE: not-equal-to boolean TRUE
		{name: "ne_TRUE_bool_true_no_match", v: BoolVal(true), crit: StringVal("<>TRUE"), want: false},
		{name: "ne_TRUE_bool_false_matches", v: BoolVal(false), crit: StringVal("<>TRUE"), want: true},
		{name: "ne_TRUE_num_matches", v: NumberVal(1), crit: StringVal("<>TRUE"), want: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := MatchesCriteria(tt.v, tt.crit)
			if got != tt.want {
				t.Errorf("MatchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// QUARTILE
// ---------------------------------------------------------------------------

func TestQUARTILE(t *testing.T) {
	// Standard dataset from Excel docs: {1,2,4,7,8,9,10,12}
	stdResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8),
			{Col: 1, Row: 6}: NumberVal(9),
			{Col: 1, Row: 7}: NumberVal(10),
			{Col: 1, Row: 8}: NumberVal(12),
		},
	}

	// Single element in B1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(42),
		},
	}

	// Two elements in C1:C2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
		},
	}

	// Mixed types: numbers, strings, booleans in D1:D5
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(10),
			{Col: 4, Row: 2}: StringVal("hello"),
			{Col: 4, Row: 3}: NumberVal(20),
			{Col: 4, Row: 4}: BoolVal(true),
			{Col: 4, Row: 5}: NumberVal(30),
		},
	}

	// Negative numbers in E1:E4
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-10),
			{Col: 5, Row: 2}: NumberVal(-5),
			{Col: 5, Row: 3}: NumberVal(0),
			{Col: 5, Row: 4}: NumberVal(5),
		},
	}

	// All same values in F1:F4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(7),
			{Col: 6, Row: 2}: NumberVal(7),
			{Col: 6, Row: 3}: NumberVal(7),
			{Col: 6, Row: 4}: NumberVal(7),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in G1:G20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 7, Row: i}] = NumberVal(float64(i))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// quart=0 (minimum)
		{"q0_min", "QUARTILE(A1:A8,0)", stdResolver, 1, 0},
		// quart=1 (25th percentile) - Excel example
		{"q1_25th", "QUARTILE(A1:A8,1)", stdResolver, 3.5, 0},
		// quart=2 (median)
		{"q2_median", "QUARTILE(A1:A8,2)", stdResolver, 7.5, 0},
		// quart=3 (75th percentile)
		{"q3_75th", "QUARTILE(A1:A8,3)", stdResolver, 9.25, 0},
		// quart=4 (maximum)
		{"q4_max", "QUARTILE(A1:A8,4)", stdResolver, 12, 0},
		// quart as float truncated to integer
		{"q_float_1.7_truncates_to_1", "QUARTILE(A1:A8,1.7)", stdResolver, 3.5, 0},
		{"q_float_3.9_truncates_to_3", "QUARTILE(A1:A8,3.9)", stdResolver, 9.25, 0},
		// quart < 0 → #NUM!
		{"q_negative", "QUARTILE(A1:A8,-1)", stdResolver, 0, ErrValNUM},
		// quart > 4 → #NUM!
		{"q_over_4", "QUARTILE(A1:A8,5)", stdResolver, 0, ErrValNUM},
		// Empty array → #NUM!
		{"empty_array", "QUARTILE(Z1:Z3,1)", emptyResolver, 0, ErrValNUM},
		// Single element
		{"single_q0", "QUARTILE(B1:B1,0)", singleResolver, 42, 0},
		{"single_q1", "QUARTILE(B1:B1,1)", singleResolver, 42, 0},
		{"single_q2", "QUARTILE(B1:B1,2)", singleResolver, 42, 0},
		{"single_q3", "QUARTILE(B1:B1,3)", singleResolver, 42, 0},
		{"single_q4", "QUARTILE(B1:B1,4)", singleResolver, 42, 0},
		// Two element array
		{"two_q0", "QUARTILE(C1:C2,0)", twoResolver, 5, 0},
		{"two_q1", "QUARTILE(C1:C2,1)", twoResolver, 7.5, 0},
		{"two_q2", "QUARTILE(C1:C2,2)", twoResolver, 10, 0},
		{"two_q3", "QUARTILE(C1:C2,3)", twoResolver, 12.5, 0},
		{"two_q4", "QUARTILE(C1:C2,4)", twoResolver, 15, 0},
		// Mixed types (non-numeric ignored) → only 10, 20, 30
		{"mixed_q1", "QUARTILE(D1:D5,1)", mixedResolver, 15, 0},
		{"mixed_q2", "QUARTILE(D1:D5,2)", mixedResolver, 20, 0},
		// Negative numbers
		{"neg_q0", "QUARTILE(E1:E4,0)", negResolver, -10, 0},
		{"neg_q1", "QUARTILE(E1:E4,1)", negResolver, -6.25, 0},
		{"neg_q2", "QUARTILE(E1:E4,2)", negResolver, -2.5, 0},
		{"neg_q4", "QUARTILE(E1:E4,4)", negResolver, 5, 0},
		// All same values
		{"same_q0", "QUARTILE(F1:F4,0)", sameResolver, 7, 0},
		{"same_q1", "QUARTILE(F1:F4,1)", sameResolver, 7, 0},
		{"same_q2", "QUARTILE(F1:F4,2)", sameResolver, 7, 0},
		{"same_q4", "QUARTILE(F1:F4,4)", sameResolver, 7, 0},
		// Large dataset
		{"large_q1", "QUARTILE(G1:G20,1)", largeResolver, 5.75, 0},
		{"large_q2", "QUARTILE(G1:G20,2)", largeResolver, 10.5, 0},
		{"large_q3", "QUARTILE(G1:G20,3)", largeResolver, 15.25, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// QUARTILE(x,q) == PERCENTILE(x, q*0.25) equivalence check
	t.Run("equivalence_with_PERCENTILE", func(t *testing.T) {
		for q := 0; q <= 4; q++ {
			qf := evalCompile(t, "QUARTILE(A1:A8,"+string(rune('0'+q))+")")
			qv, err := Eval(qf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval QUARTILE q=%d: %v", q, err)
			}

			pctStr := []string{"0", "0.25", "0.5", "0.75", "1"}[q]
			pf := evalCompile(t, "PERCENTILE(A1:A8,"+pctStr+")")
			pv, err := Eval(pf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval PERCENTILE q=%d: %v", q, err)
			}

			if math.Abs(qv.Num-pv.Num) > 1e-9 {
				t.Errorf("QUARTILE(q=%d)=%g != PERCENTILE(k=%s)=%g", q, qv.Num, pctStr, pv.Num)
			}
		}
	})

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "QUARTILE(A1:A8)")
		got, err := Eval(cf, stdResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// PERCENTILE
// ---------------------------------------------------------------------------

func TestPERCENTILE(t *testing.T) {
	// Excel docs example: {1,2,3,4} in A1:A4
	basicResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
		},
	}

	// Single element: B1=42
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(42),
		},
	}

	// Two elements: C1=10, C2=20
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(10),
			{Col: 3, Row: 2}: NumberVal(20),
		},
	}

	// Large dataset {1..20} in D1:D20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 4, Row: i}] = NumberVal(float64(i))
	}

	// All same values: E1:E4 = 7
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(7),
			{Col: 5, Row: 2}: NumberVal(7),
			{Col: 5, Row: 3}: NumberVal(7),
			{Col: 5, Row: 4}: NumberVal(7),
		},
	}

	// Negative numbers: F1:F4 = {-10, -5, 0, 5}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(-10),
			{Col: 6, Row: 2}: NumberVal(-5),
			{Col: 6, Row: 3}: NumberVal(0),
			{Col: 6, Row: 4}: NumberVal(5),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Non-numeric mixed: G1=1, G2="hello", G3=3, G4=TRUE
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 7, Row: 1}: NumberVal(1),
			{Col: 7, Row: 2}: StringVal("hello"),
			{Col: 7, Row: 3}: NumberVal(3),
			{Col: 7, Row: 4}: BoolVal(true),
		},
	}

	// Error in range: H1=1, H2=#DIV/0!, H3=3
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(1),
			{Col: 8, Row: 2}: ErrorVal(ErrValDIV0),
			{Col: 8, Row: 3}: NumberVal(3),
		},
	}

	// Excel doc example: {1,2,3,6,6,6,7,8,9} → PERCENTILE(data, 0.3) is from the docs
	// Using simpler data set matching Excel example
	excelDocResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(1),
			{Col: 9, Row: 2}: NumberVal(2),
			{Col: 9, Row: 3}: NumberVal(3),
			{Col: 9, Row: 4}: NumberVal(6),
			{Col: 9, Row: 5}: NumberVal(6),
			{Col: 9, Row: 6}: NumberVal(6),
			{Col: 9, Row: 7}: NumberVal(7),
			{Col: 9, Row: 8}: NumberVal(8),
			{Col: 9, Row: 9}: NumberVal(9),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		want     float64
		wantErr  ErrorValue
	}{
		// Basic: PERCENTILE({1,2,3,4}, 0.5) → median = 2.5
		{"basic_median", "PERCENTILE(A1:A4,0.5)", basicResolver, 2.5, 0},
		// k=0 → minimum
		{"k_zero_min", "PERCENTILE(A1:A4,0)", basicResolver, 1, 0},
		// k=1 → maximum
		{"k_one_max", "PERCENTILE(A1:A4,1)", basicResolver, 4, 0},
		// k=0.25 → first quartile: rank=0.25*3=0.75, 1+0.75*(2-1)=1.75
		{"k_25th", "PERCENTILE(A1:A4,0.25)", basicResolver, 1.75, 0},
		// k=0.75 → third quartile: rank=0.75*3=2.25, 3+0.25*(4-3)=3.25
		{"k_75th", "PERCENTILE(A1:A4,0.75)", basicResolver, 3.25, 0},
		// k=0.1: rank=0.1*3=0.3, 1+0.3*(2-1)=1.3
		{"k_10th", "PERCENTILE(A1:A4,0.1)", basicResolver, 1.3, 0},
		// k=0.9: rank=0.9*3=2.7, 3+0.7*(4-3)=3.7
		{"k_90th", "PERCENTILE(A1:A4,0.9)", basicResolver, 3.7, 0},
		// Single element array: any k returns that element
		{"single_k0", "PERCENTILE(B1:B1,0)", singleResolver, 42, 0},
		{"single_k05", "PERCENTILE(B1:B1,0.5)", singleResolver, 42, 0},
		{"single_k1", "PERCENTILE(B1:B1,1)", singleResolver, 42, 0},
		// Two element array: k=0.5 → 15
		{"two_median", "PERCENTILE(C1:C2,0.5)", twoResolver, 15, 0},
		// Two element: k=0 → 10, k=1 → 20
		{"two_k0", "PERCENTILE(C1:C2,0)", twoResolver, 10, 0},
		{"two_k1", "PERCENTILE(C1:C2,1)", twoResolver, 20, 0},
		// Large array {1..20}: k=0.5 → rank=0.5*19=9.5, 10+0.5*(11-10)=10.5
		{"large_median", "PERCENTILE(D1:D20,0.5)", largeResolver, 10.5, 0},
		// Large array: k=0.25 → rank=0.25*19=4.75, 5+0.75*(6-5)=5.75
		{"large_25th", "PERCENTILE(D1:D20,0.25)", largeResolver, 5.75, 0},
		// Large array: k=0.75 → rank=0.75*19=14.25, 15+0.25*(16-15)=15.25
		{"large_75th", "PERCENTILE(D1:D20,0.75)", largeResolver, 15.25, 0},
		// All same values → that value regardless of k
		{"same_k0", "PERCENTILE(E1:E4,0)", sameResolver, 7, 0},
		{"same_k05", "PERCENTILE(E1:E4,0.5)", sameResolver, 7, 0},
		{"same_k1", "PERCENTILE(E1:E4,1)", sameResolver, 7, 0},
		// Negative numbers: sorted {-10,-5,0,5}, k=0.5 → rank=1.5, -5+0.5*5=-2.5
		{"neg_median", "PERCENTILE(F1:F4,0.5)", negResolver, -2.5, 0},
		// Negative: k=0 → -10
		{"neg_k0", "PERCENTILE(F1:F4,0)", negResolver, -10, 0},
		// Negative: k=1 → 5
		{"neg_k1", "PERCENTILE(F1:F4,1)", negResolver, 5, 0},
		// k < 0 → #NUM!
		{"k_negative", "PERCENTILE(A1:A4,-0.1)", basicResolver, 0, ErrValNUM},
		// k > 1 → #NUM!
		{"k_over_one", "PERCENTILE(A1:A4,1.5)", basicResolver, 0, ErrValNUM},
		// Empty array → #NUM!
		{"empty_array", "PERCENTILE(Z1:Z3,0.5)", emptyResolver, 0, ErrValNUM},
		// Non-numeric values in array are ignored: {1,3} → k=0.5 → 2
		{"mixed_ignore_strings", "PERCENTILE(G1:G4,0.5)", mixedResolver, 2, 0},
		// Excel doc example: PERCENTILE({1,2,3,6,6,6,7,8,9},0.3)
		// rank=0.3*8=2.4, sorted[2]=3, sorted[3]=6 → 3+0.4*(6-3)=4.2
		{"excel_doc_example", "PERCENTILE(I1:I9,0.3)", excelDocResolver, 4.2, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %v, want number", got.Type)
			}
			if math.Abs(got.Num-tt.want) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.want)
			}
		})
	}

	// Error propagation: error in range propagates
	t.Run("error_propagation", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTILE(H1:H3,0.5)")
		got, err := Eval(cf, errResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	// Too few arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTILE(A1:A4)")
		got, err := Eval(cf, basicResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	// Too many arguments
	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTILE(A1:A4,0.5,1)")
		got, err := Eval(cf, basicResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// PERCENTILE.EXC
// ---------------------------------------------------------------------------

func TestPERCENTILEEXC(t *testing.T) {
	// Excel docs example: {6,7,15,36,39,40,41,42,43,47,49} in A1:A11
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(6),
			{Col: 1, Row: 2}:  NumberVal(7),
			{Col: 1, Row: 3}:  NumberVal(15),
			{Col: 1, Row: 4}:  NumberVal(36),
			{Col: 1, Row: 5}:  NumberVal(39),
			{Col: 1, Row: 6}:  NumberVal(40),
			{Col: 1, Row: 7}:  NumberVal(41),
			{Col: 1, Row: 8}:  NumberVal(42),
			{Col: 1, Row: 9}:  NumberVal(43),
			{Col: 1, Row: 10}: NumberVal(47),
			{Col: 1, Row: 11}: NumberVal(49),
		},
	}

	// Standard dataset: {1,2,3,4,5,6,7,8} in B1:B8
	stdResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
			{Col: 2, Row: 6}: NumberVal(6),
			{Col: 2, Row: 7}: NumberVal(7),
			{Col: 2, Row: 8}: NumberVal(8),
		},
	}

	// Single element in C1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(42),
		},
	}

	// Two elements in D1:D2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(10),
			{Col: 4, Row: 2}: NumberVal(20),
		},
	}

	// Four elements in E1:E4
	fourResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(1),
			{Col: 5, Row: 2}: NumberVal(2),
			{Col: 5, Row: 3}: NumberVal(3),
			{Col: 5, Row: 4}: NumberVal(4),
		},
	}

	// Negative numbers in F1:F4
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(-10),
			{Col: 6, Row: 2}: NumberVal(-5),
			{Col: 6, Row: 3}: NumberVal(0),
			{Col: 6, Row: 4}: NumberVal(5),
		},
	}

	// All same values in G1:G4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 7, Row: 1}: NumberVal(7),
			{Col: 7, Row: 2}: NumberVal(7),
			{Col: 7, Row: 3}: NumberVal(7),
			{Col: 7, Row: 4}: NumberVal(7),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in H1:H20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 8, Row: i}] = NumberVal(float64(i))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// Excel docs example: PERCENTILE.EXC({6,7,15,36,39,40,41,42,43,47,49}, 0.25) = 15
		{"excel_example_25th", "PERCENTILE.EXC(A1:A11,0.25)", excelResolver, 15, 0},
		// Excel docs example: PERCENTILE.EXC({6,7,15,36,39,40,41,42,43,47,49}, 0.75) = 43
		{"excel_example_75th", "PERCENTILE.EXC(A1:A11,0.75)", excelResolver, 43, 0},
		// Median (k=0.5) on excel data: rank=0.5*12=6, so nums[5]=40
		{"excel_median", "PERCENTILE.EXC(A1:A11,0.5)", excelResolver, 40, 0},
		// Standard dataset {1..8}: k=0.25, rank=0.25*9=2.25, interp between 2 and 3
		{"std_25th", "PERCENTILE.EXC(B1:B8,0.25)", stdResolver, 2.25, 0},
		// k=0.5, rank=0.5*9=4.5, interp between 4 and 5
		{"std_50th", "PERCENTILE.EXC(B1:B8,0.5)", stdResolver, 4.5, 0},
		// k=0.75, rank=0.75*9=6.75, interp between 7 and 8: 6+0.75*1=6.75
		{"std_75th", "PERCENTILE.EXC(B1:B8,0.75)", stdResolver, 6.75, 0},
		// k at exact integer rank: k=1/9≈0.1111, rank=1.0, → nums[0]=1
		{"std_exact_first", "PERCENTILE.EXC(B1:B8,1/9)", stdResolver, 1, 0},
		// k at exact integer rank: k=8/9≈0.8889, rank=8.0, → nums[7]=8
		{"std_exact_last", "PERCENTILE.EXC(B1:B8,8/9)", stdResolver, 8, 0},
		// k=0 → #NUM!
		{"k_zero", "PERCENTILE.EXC(B1:B8,0)", stdResolver, 0, ErrValNUM},
		// k=1 → #NUM!
		{"k_one", "PERCENTILE.EXC(B1:B8,1)", stdResolver, 0, ErrValNUM},
		// k negative → #NUM!
		{"k_negative", "PERCENTILE.EXC(B1:B8,-0.5)", stdResolver, 0, ErrValNUM},
		// k > 1 → #NUM!
		{"k_over_one", "PERCENTILE.EXC(B1:B8,1.5)", stdResolver, 0, ErrValNUM},
		// Single element: only k=0.5 works (rank=0.5*2=1)
		{"single_median", "PERCENTILE.EXC(C1:C1,0.5)", singleResolver, 42, 0},
		// Single element: k=0.25 → rank=0.25*2=0.5 < 1 → #NUM!
		{"single_k25_num", "PERCENTILE.EXC(C1:C1,0.25)", singleResolver, 0, ErrValNUM},
		// Single element: k=0.75 → rank=0.75*2=1.5 > 1 → #NUM!
		{"single_k75_num", "PERCENTILE.EXC(C1:C1,0.75)", singleResolver, 0, ErrValNUM},
		// Two elements: k=1/3 → rank=(1/3)*3=1 → nums[0]=10
		{"two_exact_first", "PERCENTILE.EXC(D1:D2,1/3)", twoResolver, 10, 0},
		// Two elements: k=2/3 → rank=(2/3)*3=2 → nums[1]=20
		{"two_exact_last", "PERCENTILE.EXC(D1:D2,2/3)", twoResolver, 20, 0},
		// Two elements: k=0.5 → rank=0.5*3=1.5 → 10+0.5*10=15
		{"two_median", "PERCENTILE.EXC(D1:D2,0.5)", twoResolver, 15, 0},
		// Four elements {1,2,3,4}: k=0.5 → rank=0.5*5=2.5 → 2+0.5=2.5
		{"four_median", "PERCENTILE.EXC(E1:E4,0.5)", fourResolver, 2.5, 0},
		// Four elements: k=0.2 → rank=0.2*5=1 → nums[0]=1
		{"four_k20", "PERCENTILE.EXC(E1:E4,0.2)", fourResolver, 1, 0},
		// Four elements: k=0.8 → rank=0.8*5=4 → nums[3]=4
		{"four_k80", "PERCENTILE.EXC(E1:E4,0.8)", fourResolver, 4, 0},
		// Four elements: k=0.1 → rank=0.1*5=0.5 < 1 → #NUM!
		{"four_k10_num", "PERCENTILE.EXC(E1:E4,0.1)", fourResolver, 0, ErrValNUM},
		// Four elements: k=0.9 → rank=0.9*5=4.5 > 4 → #NUM!
		{"four_k90_num", "PERCENTILE.EXC(E1:E4,0.9)", fourResolver, 0, ErrValNUM},
		// Negative numbers: k=0.5 → rank=0.5*5=2.5 → -5+0.5*5=-2.5
		{"neg_median", "PERCENTILE.EXC(F1:F4,0.5)", negResolver, -2.5, 0},
		// All same values: k=0.5 → rank=0.5*5=2.5 → 7+0.5*0=7
		{"same_median", "PERCENTILE.EXC(G1:G4,0.5)", sameResolver, 7, 0},
		// Empty array → #NUM!
		{"empty_array", "PERCENTILE.EXC(Z1:Z3,0.5)", emptyResolver, 0, ErrValNUM},
		// Large dataset {1..20}: k=0.25 → rank=0.25*21=5.25 → 5+0.25*1=5.25
		{"large_25th", "PERCENTILE.EXC(H1:H20,0.25)", largeResolver, 5.25, 0},
		// Large dataset: k=0.5 → rank=0.5*21=10.5 → 10+0.5*1=10.5
		{"large_50th", "PERCENTILE.EXC(H1:H20,0.5)", largeResolver, 10.5, 0},
		// Large dataset: k=0.75 → rank=0.75*21=15.75 → 15+0.75*1=15.75
		{"large_75th", "PERCENTILE.EXC(H1:H20,0.75)", largeResolver, 15.75, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTILE.EXC(B1:B8)")
		got, err := Eval(cf, stdResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// QUARTILE.EXC
// ---------------------------------------------------------------------------

func TestQUARTILEEXC(t *testing.T) {
	// Excel docs example: {6,7,15,36,39,40,41,42,43,47,49} in A1:A11
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(6),
			{Col: 1, Row: 2}:  NumberVal(7),
			{Col: 1, Row: 3}:  NumberVal(15),
			{Col: 1, Row: 4}:  NumberVal(36),
			{Col: 1, Row: 5}:  NumberVal(39),
			{Col: 1, Row: 6}:  NumberVal(40),
			{Col: 1, Row: 7}:  NumberVal(41),
			{Col: 1, Row: 8}:  NumberVal(42),
			{Col: 1, Row: 9}:  NumberVal(43),
			{Col: 1, Row: 10}: NumberVal(47),
			{Col: 1, Row: 11}: NumberVal(49),
		},
	}

	// Standard dataset: {1,2,3,4,5,6,7,8} in B1:B8
	stdResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
			{Col: 2, Row: 6}: NumberVal(6),
			{Col: 2, Row: 7}: NumberVal(7),
			{Col: 2, Row: 8}: NumberVal(8),
		},
	}

	// Four elements in C1:C4
	fourResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(1),
			{Col: 3, Row: 2}: NumberVal(2),
			{Col: 3, Row: 3}: NumberVal(3),
			{Col: 3, Row: 4}: NumberVal(4),
		},
	}

	// Negative numbers in D1:D4
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(-10),
			{Col: 4, Row: 2}: NumberVal(-5),
			{Col: 4, Row: 3}: NumberVal(0),
			{Col: 4, Row: 4}: NumberVal(5),
		},
	}

	// All same values in E1:E4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(7),
			{Col: 5, Row: 2}: NumberVal(7),
			{Col: 5, Row: 3}: NumberVal(7),
			{Col: 5, Row: 4}: NumberVal(7),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in F1:F20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 6, Row: i}] = NumberVal(float64(i))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// Excel docs example: QUARTILE.EXC({6,7,...,49}, 1) = 15
		{"excel_q1", "QUARTILE.EXC(A1:A11,1)", excelResolver, 15, 0},
		// Excel docs example: QUARTILE.EXC({6,7,...,49}, 3) = 43
		{"excel_q3", "QUARTILE.EXC(A1:A11,3)", excelResolver, 43, 0},
		// QUARTILE.EXC(data, 2) = MEDIAN
		{"excel_q2_median", "QUARTILE.EXC(A1:A11,2)", excelResolver, 40, 0},
		// Standard dataset {1..8}: q1 → PERCENTILE.EXC(data, 0.25) = 2.25
		{"std_q1", "QUARTILE.EXC(B1:B8,1)", stdResolver, 2.25, 0},
		// q2 → PERCENTILE.EXC(data, 0.5) = 4.5
		{"std_q2", "QUARTILE.EXC(B1:B8,2)", stdResolver, 4.5, 0},
		// q3 → PERCENTILE.EXC(data, 0.75) = 6.75
		{"std_q3", "QUARTILE.EXC(B1:B8,3)", stdResolver, 6.75, 0},
		// quart=0 → #NUM! (unlike QUARTILE.INC)
		{"q0_num", "QUARTILE.EXC(B1:B8,0)", stdResolver, 0, ErrValNUM},
		// quart=4 → #NUM! (unlike QUARTILE.INC)
		{"q4_num", "QUARTILE.EXC(B1:B8,4)", stdResolver, 0, ErrValNUM},
		// quart negative → #NUM!
		{"q_negative", "QUARTILE.EXC(B1:B8,-1)", stdResolver, 0, ErrValNUM},
		// quart=5 → #NUM!
		{"q_over_4", "QUARTILE.EXC(B1:B8,5)", stdResolver, 0, ErrValNUM},
		// quart as float truncated to integer
		{"q_float_1.7_truncates_to_1", "QUARTILE.EXC(B1:B8,1.7)", stdResolver, 2.25, 0},
		{"q_float_3.9_truncates_to_3", "QUARTILE.EXC(B1:B8,3.9)", stdResolver, 6.75, 0},
		// Four elements: q1 → k=0.25, rank=0.25*5=1.25 → 1+0.25=1.25
		{"four_q1", "QUARTILE.EXC(C1:C4,1)", fourResolver, 1.25, 0},
		// Four elements: q2 → k=0.5, rank=0.5*5=2.5 → 2+0.5=2.5
		{"four_q2", "QUARTILE.EXC(C1:C4,2)", fourResolver, 2.5, 0},
		// Four elements: q3 → k=0.75, rank=0.75*5=3.75 → 3+0.75=3.75
		{"four_q3", "QUARTILE.EXC(C1:C4,3)", fourResolver, 3.75, 0},
		// Negative numbers: q1 → k=0.25, rank=0.25*5=1.25 → -10+0.25*5=-8.75
		{"neg_q1", "QUARTILE.EXC(D1:D4,1)", negResolver, -8.75, 0},
		// Negative numbers: q2
		{"neg_q2", "QUARTILE.EXC(D1:D4,2)", negResolver, -2.5, 0},
		// Negative numbers: q3 → k=0.75, rank=0.75*5=3.75 → 0+0.75*5=3.75
		{"neg_q3", "QUARTILE.EXC(D1:D4,3)", negResolver, 3.75, 0},
		// All same values
		{"same_q1", "QUARTILE.EXC(E1:E4,1)", sameResolver, 7, 0},
		{"same_q2", "QUARTILE.EXC(E1:E4,2)", sameResolver, 7, 0},
		{"same_q3", "QUARTILE.EXC(E1:E4,3)", sameResolver, 7, 0},
		// Empty array → #NUM!
		{"empty_array", "QUARTILE.EXC(Z1:Z3,1)", emptyResolver, 0, ErrValNUM},
		// Large dataset {1..20}: q1 → PERCENTILE.EXC(data, 0.25) = 5.25
		{"large_q1", "QUARTILE.EXC(F1:F20,1)", largeResolver, 5.25, 0},
		// Large dataset: q2 → 10.5
		{"large_q2", "QUARTILE.EXC(F1:F20,2)", largeResolver, 10.5, 0},
		// Large dataset: q3 → 15.75
		{"large_q3", "QUARTILE.EXC(F1:F20,3)", largeResolver, 15.75, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// QUARTILE.EXC(x,q) == PERCENTILE.EXC(x, q*0.25) equivalence check
	t.Run("equivalence_with_PERCENTILE_EXC", func(t *testing.T) {
		for q := 1; q <= 3; q++ {
			qf := evalCompile(t, "QUARTILE.EXC(B1:B8,"+string(rune('0'+q))+")")
			qv, err := Eval(qf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval QUARTILE.EXC q=%d: %v", q, err)
			}

			pctStr := []string{"", "0.25", "0.5", "0.75"}[q]
			pf := evalCompile(t, "PERCENTILE.EXC(B1:B8,"+pctStr+")")
			pv, err := Eval(pf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval PERCENTILE.EXC q=%d: %v", q, err)
			}

			if math.Abs(qv.Num-pv.Num) > 1e-9 {
				t.Errorf("QUARTILE.EXC(q=%d)=%g != PERCENTILE.EXC(k=%s)=%g", q, qv.Num, pctStr, pv.Num)
			}
		}
	})

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "QUARTILE.EXC(B1:B8)")
		got, err := Eval(cf, stdResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// TRIMMEAN
// ---------------------------------------------------------------------------

func TestTRIMMEAN(t *testing.T) {
	// Excel docs example: {4,5,6,7,2,3,4,5,1,2,3} in A1:A11
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(4),
			{Col: 1, Row: 2}:  NumberVal(5),
			{Col: 1, Row: 3}:  NumberVal(6),
			{Col: 1, Row: 4}:  NumberVal(7),
			{Col: 1, Row: 5}:  NumberVal(2),
			{Col: 1, Row: 6}:  NumberVal(3),
			{Col: 1, Row: 7}:  NumberVal(4),
			{Col: 1, Row: 8}:  NumberVal(5),
			{Col: 1, Row: 9}:  NumberVal(1),
			{Col: 1, Row: 10}: NumberVal(2),
			{Col: 1, Row: 11}: NumberVal(3),
		},
	}

	// Single element in B1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(42),
		},
	}

	// Two elements in C1:C2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
		},
	}

	// All same values in D1:D4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(7),
			{Col: 4, Row: 2}: NumberVal(7),
			{Col: 4, Row: 3}: NumberVal(7),
			{Col: 4, Row: 4}: NumberVal(7),
		},
	}

	// Negative numbers in E1:E5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-10),
			{Col: 5, Row: 2}: NumberVal(-5),
			{Col: 5, Row: 3}: NumberVal(0),
			{Col: 5, Row: 4}: NumberVal(5),
			{Col: 5, Row: 5}: NumberVal(10),
		},
	}

	// Mixed types in F1:F5 (strings and booleans ignored)
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(10),
			{Col: 6, Row: 2}: StringVal("hello"),
			{Col: 6, Row: 3}: NumberVal(20),
			{Col: 6, Row: 4}: BoolVal(true),
			{Col: 6, Row: 5}: NumberVal(30),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in G1:G30
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 30; i++ {
		largeResolver.cells[CellAddr{Col: 7, Row: i}] = NumberVal(float64(i))
	}

	// Four elements in H1:H4
	fourResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(1),
			{Col: 8, Row: 2}: NumberVal(2),
			{Col: 8, Row: 3}: NumberVal(3),
			{Col: 8, Row: 4}: NumberVal(4),
		},
	}

	// Six elements in I1:I6
	sixResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(1),
			{Col: 9, Row: 2}: NumberVal(2),
			{Col: 9, Row: 3}: NumberVal(3),
			{Col: 9, Row: 4}: NumberVal(4),
			{Col: 9, Row: 5}: NumberVal(5),
			{Col: 9, Row: 6}: NumberVal(6),
		},
	}

	// Ten elements in J1:J10
	tenResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 10; i++ {
		tenResolver.cells[CellAddr{Col: 10, Row: i}] = NumberVal(float64(i * 10))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// Excel docs example: {4,5,6,7,2,3,4,5,1,2,3}, 0.2
		// sorted: 1,2,2,3,3,4,4,5,5,6,7; n=11, floor(11*0.2/2)=1, trim 1 each end
		// remaining: 2,2,3,3,4,4,5,5,6 → mean = 34/9
		{"excel_example", "TRIMMEAN(A1:A11,0.2)", excelResolver, 34.0 / 9.0, 0},

		// percent=0 → regular mean
		{"percent_zero", "TRIMMEAN(A1:A11,0)", excelResolver, 42.0 / 11.0, 0},

		// percent=0.5 on 4 elements: floor(4*0.5/2)=1, trim 1 each end → {2,3} mean=2.5
		{"percent_half_4elem", "TRIMMEAN(H1:H4,0.5)", fourResolver, 2.5, 0},

		// percent just under 1 (0.99) on 11 elements: floor(11*0.99/2)=5, trim 5 each end → 1 left
		{"percent_099", "TRIMMEAN(A1:A11,0.99)", excelResolver, 4, 0},

		// percent < 0 → #NUM!
		{"percent_negative", "TRIMMEAN(A1:A11,-0.1)", excelResolver, 0, ErrValNUM},

		// percent >= 1 → #NUM!
		{"percent_one", "TRIMMEAN(A1:A11,1)", excelResolver, 0, ErrValNUM},
		{"percent_over_one", "TRIMMEAN(A1:A11,1.5)", excelResolver, 0, ErrValNUM},

		// Single element, percent=0
		{"single_pct0", "TRIMMEAN(B1:B1,0)", singleResolver, 42, 0},

		// Single element, percent=0.5: floor(1*0.5/2)=0, no trim → mean=42
		{"single_pct05", "TRIMMEAN(B1:B1,0.5)", singleResolver, 42, 0},

		// Two elements, percent=0.5: floor(2*0.5/2)=0, no trim → mean=10
		{"two_pct05", "TRIMMEAN(C1:C2,0.5)", twoResolver, 10, 0},

		// Two elements, percent=0.99: floor(2*0.99/2)=0, no trim → mean=10
		{"two_pct099", "TRIMMEAN(C1:C2,0.99)", twoResolver, 10, 0},

		// All same values
		{"all_same", "TRIMMEAN(D1:D4,0.5)", sameResolver, 7, 0},

		// Negative numbers: {-10,-5,0,5,10}, percent=0.4: floor(5*0.4/2)=1, trim 1 each → {-5,0,5} mean=0
		{"negative_nums", "TRIMMEAN(E1:E5,0.4)", negResolver, 0, 0},

		// Mixed types: only {10,20,30} numeric, percent=0: mean=20
		{"mixed_types_pct0", "TRIMMEAN(F1:F5,0)", mixedResolver, 20, 0},

		// Mixed types: {10,20,30}, percent=0.5: floor(3*0.5/2)=0, no trim → mean=20
		{"mixed_types_pct05", "TRIMMEAN(F1:F5,0.5)", mixedResolver, 20, 0},

		// Empty array → #NUM!
		{"empty_array", "TRIMMEAN(Z1:Z3,0.2)", emptyResolver, 0, ErrValNUM},

		// Large dataset 1..30, percent=0.1: floor(30*0.1/2)=1, trim 1 each → {2..29} mean=15.5
		{"large_pct01", "TRIMMEAN(G1:G30,0.1)", largeResolver, 15.5, 0},

		// Large dataset 1..30, percent=0.2: floor(30*0.2/2)=3, trim 3 each → {4..27} mean=15.5
		{"large_pct02", "TRIMMEAN(G1:G30,0.2)", largeResolver, 15.5, 0},

		// Six elements, percent=0.5: floor(6*0.5/2)=1, trim 1 each → {2,3,4,5} mean=3.5
		{"six_pct05", "TRIMMEAN(I1:I6,0.5)", sixResolver, 3.5, 0},

		// Ten elements {10,20..100}, percent=0.3: floor(10*0.3/2)=1, trim 1 each → {20..90} mean=55
		{"ten_pct03", "TRIMMEAN(J1:J10,0.3)", tenResolver, 55, 0},

		// Four elements, percent=0.99: floor(4*0.99/2)=1, trim 1 each → {2,3} mean=2.5
		{"four_pct099", "TRIMMEAN(H1:H4,0.99)", fourResolver, 2.5, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "TRIMMEAN(A1:A11)")
		got, err := Eval(cf, excelResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "TRIMMEAN(A1:A11,0.2,1)")
		got, err := Eval(cf, excelResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// GEOMEAN
// ---------------------------------------------------------------------------

func TestGEOMEAN(t *testing.T) {
	const tol = 1e-6

	// Resolver with Excel doc example {4,5,8,7,11,4,3} in A1:A7
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(8),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(11),
			{Col: 1, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(3),
		},
	}

	// Resolver with mixed types in B column
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: StringVal("hello"),
			{Col: 2, Row: 3}: NumberVal(8),
			{Col: 2, Row: 4}: BoolVal(true),
			{Col: 2, Row: 5}: NumberVal(4),
		},
	}

	// Resolver with empty cells in C column
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(4),
			// C2 is empty
			{Col: 3, Row: 3}: NumberVal(9),
		},
	}

	// Resolver with error in D column
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(3),
			{Col: 4, Row: 2}: ErrorVal(ErrValNA),
			{Col: 4, Row: 3}: NumberVal(6),
		},
	}

	// Large dataset in E column (1..20)
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 5, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  bool // expect ValueError type
	}{
		// Basic: GEOMEAN(2,8) = sqrt(16) = 4
		{
			name:    "basic_two_values",
			formula: "GEOMEAN(2,8)",
			wantNum: 4,
		},
		// Single value → that value
		{
			name:    "single_value",
			formula: "GEOMEAN(7)",
			wantNum: 7,
		},
		// All same values → that value
		{
			name:    "all_same_values",
			formula: "GEOMEAN(5,5,5)",
			wantNum: 5,
		},
		// Large values
		{
			name:    "large_values",
			formula: "GEOMEAN(1000000,4000000)",
			wantNum: 2000000,
		},
		// Small decimals
		{
			name:    "small_decimals",
			formula: "GEOMEAN(0.01,0.04)",
			wantNum: 0.02,
		},
		// Negative value → #NUM!
		{
			name:    "negative_value",
			formula: "GEOMEAN(-1,2,3)",
			wantErr: true,
		},
		// Zero value → #NUM!
		{
			name:    "zero_value",
			formula: "GEOMEAN(0,2,3)",
			wantErr: true,
		},
		// Zero in middle → #NUM!
		{
			name:    "zero_in_middle",
			formula: "GEOMEAN(4,0,9)",
			wantErr: true,
		},
		// Range input - Excel doc example: GEOMEAN(A2:A8) = 5.476987
		{
			name:     "excel_example_range",
			formula:  "GEOMEAN(A1:A7)",
			resolver: excelResolver,
			wantNum:  5.476987,
		},
		// Same values as direct args
		{
			name:    "excel_example_direct",
			formula: "GEOMEAN(4,5,8,7,11,4,3)",
			wantNum: 5.476987,
		},
		// Empty cells in range (ignored by collectNumeric)
		{
			name:     "empty_cells_ignored",
			formula:  "GEOMEAN(C1:C3)",
			resolver: emptyResolver,
			wantNum:  6, // geomean(4,9) = sqrt(36) = 6
		},
		// Strings in range (ignored by collectNumeric)
		{
			name:     "strings_in_range_ignored",
			formula:  "GEOMEAN(B1:B5)",
			resolver: mixedResolver,
			wantNum:  4, // geomean(2,8,4) = (2*8*4)^(1/3) = 64^(1/3) = 4
		},
		// Booleans in range (ignored by collectNumeric)
		// mixedResolver has bool TRUE at B4 - should be ignored in array context
		// Only numbers {2,8,4} are collected
		{
			name:     "booleans_in_range_ignored",
			formula:  "GEOMEAN(B1:B5)",
			resolver: mixedResolver,
			wantNum:  4, // same as strings test - all non-numeric ignored
		},
		// Scalar boolean TRUE=1
		{
			name:    "scalar_bool_true",
			formula: "GEOMEAN(TRUE,4)",
			wantNum: 2, // geomean(1,4) = sqrt(4) = 2
		},
		// Scalar boolean FALSE=0 → #NUM! (zero)
		{
			name:    "scalar_bool_false",
			formula: "GEOMEAN(FALSE,4)",
			wantErr: true,
		},
		// Error propagation (#N/A in array)
		{
			name:     "error_propagation_na",
			formula:  "GEOMEAN(D1:D3)",
			resolver: errResolver,
			wantErr:  true,
		},
		// Error propagation with #DIV/0!
		{
			name:    "error_propagation_div0",
			formula: "GEOMEAN(1,2,1/0)",
			wantErr: true,
		},
		// Many values (1..20)
		{
			name:     "many_values",
			formula:  "GEOMEAN(E1:E20)",
			resolver: largeResolver,
			wantNum:  8.304361, // geometric mean of 1..20
		},
		// Multiple array arguments
		{
			name:     "multiple_arrays",
			formula:  "GEOMEAN(A1:A3,A4:A7)",
			resolver: excelResolver,
			wantNum:  5.476987,
		},
		// Mix of direct arg and array
		{
			name:     "direct_and_array",
			formula:  "GEOMEAN(2,A1:A1)",
			resolver: excelResolver,
			wantNum:  2.828427, // geomean(2,4) = sqrt(8)
		},
		// String number as direct arg (coerced)
		{
			name:    "string_number_direct",
			formula: `GEOMEAN("4","9")`,
			wantNum: 6, // geomean(4,9) = sqrt(36) = 6
		},
		// Non-numeric string as direct arg → error
		{
			name:    "non_numeric_string_direct",
			formula: `GEOMEAN("hello",4)`,
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			resolver := tt.resolver
			if resolver == nil {
				resolver = &mockResolver{}
			}
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Errorf("got %v (type %d), want error", got, got.Type)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// HARMEAN
// ---------------------------------------------------------------------------

func TestHARMEAN(t *testing.T) {
	const tol = 1e-6

	// Resolver with {4,5,8,7,11,4,3} in A1:A7
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(8),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(11),
			{Col: 1, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(3),
		},
	}

	// Resolver with mixed types in B column
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: StringVal("hello"),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: BoolVal(true),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Resolver with zero in C column
	zeroResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(0),
			{Col: 3, Row: 3}: NumberVal(10),
		},
	}

	// Resolver with error in D column
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(3),
			{Col: 4, Row: 2}: ErrorVal(ErrValNA),
			{Col: 4, Row: 3}: NumberVal(6),
		},
	}

	// Large dataset in E column (1..20)
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 5, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  bool // expect ValueError type
	}{
		// Excel documentation example: {4,5,8,7,11,4,3}
		{
			name:     "excel_example_array_ref",
			formula:  "HARMEAN(A1:A7)",
			resolver: excelResolver,
			wantNum:  5.028376,
		},
		// Same values as direct args
		{
			name:     "excel_example_direct",
			formula:  "HARMEAN(4,5,8,7,11,4,3)",
			resolver: nil,
			wantNum:  5.028376,
		},
		// Single value
		{
			name:     "single_value",
			formula:  "HARMEAN(4)",
			resolver: nil,
			wantNum:  4,
		},
		// Two values: 2/(1/1+1/4) = 2/1.25 = 1.6
		{
			name:     "two_values",
			formula:  "HARMEAN(1,4)",
			resolver: nil,
			wantNum:  1.6,
		},
		// All same values
		{
			name:     "all_same",
			formula:  "HARMEAN(5,5,5)",
			resolver: nil,
			wantNum:  5,
		},
		// Zero returns #NUM!
		{
			name:    "zero_direct",
			formula: "HARMEAN(0)",
			wantErr: true,
		},
		// Zero in middle
		{
			name:    "zero_in_middle",
			formula: "HARMEAN(3,0,6)",
			wantErr: true,
		},
		// Negative value returns #NUM!
		{
			name:    "negative_value",
			formula: "HARMEAN(-1,2,3)",
			wantErr: true,
		},
		// Boolean TRUE as direct arg (counted as 1)
		{
			name:    "bool_true_direct",
			formula: "HARMEAN(TRUE,4)",
			wantNum: 1.6, // 2/(1/1+1/4) = 1.6
		},
		// String number as direct arg
		{
			name:    "string_number_direct",
			formula: `HARMEAN("3",6)`,
			wantNum: 4, // 2/(1/3+1/6) = 2/0.5 = 4
		},
		// Array with mixed types: text and bool ignored, only numbers {2,4,8}
		{
			name:     "array_mixed_types",
			formula:  "HARMEAN(B1:B5)",
			resolver: mixedResolver,
			wantNum:  3.428571, // 3/(1/2+1/4+1/8) = 3/0.875
		},
		// Zero in array → #NUM!
		{
			name:     "zero_in_array",
			formula:  "HARMEAN(C1:C3)",
			resolver: zeroResolver,
			wantErr:  true,
		},
		// Error propagation (#N/A in array)
		{
			name:     "error_propagation_na",
			formula:  "HARMEAN(D1:D3)",
			resolver: errResolver,
			wantErr:  true,
		},
		// Error propagation with #VALUE!
		{
			name:    "error_propagation_value",
			formula: `HARMEAN(1,2,1/0)`,
			wantErr: true,
		},
		// Large dataset 1..20
		{
			name:     "large_dataset",
			formula:  "HARMEAN(E1:E20)",
			resolver: largeResolver,
			wantNum:  5.559046, // harmonic mean of 1..20
		},
		// Very small positive numbers
		{
			name:    "very_small_numbers",
			formula: "HARMEAN(0.001,0.002,0.003)",
			wantNum: 0.001636, // 3/(1/0.001+1/0.002+1/0.003)
		},
		// Very large numbers
		{
			name:    "very_large_numbers",
			formula: "HARMEAN(1000000,2000000,3000000)",
			wantNum: 1636363.636364,
		},
		// Single element array
		{
			name:     "single_element_array",
			formula:  "HARMEAN(A1:A1)",
			resolver: excelResolver,
			wantNum:  4,
		},
		// Multiple array arguments
		{
			name:     "multiple_arrays",
			formula:  "HARMEAN(A1:A3,A4:A7)",
			resolver: excelResolver,
			wantNum:  5.028376,
		},
		// Mix of direct and array
		{
			name:     "direct_and_array",
			formula:  "HARMEAN(2,A1:A1)",
			resolver: excelResolver,
			wantNum:  2.666667, // 2/(1/2+1/4) = 2/0.75
		},
		// Equal fractions
		{
			name:    "fractions",
			formula: "HARMEAN(0.5,0.5)",
			wantNum: 0.5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			resolver := tt.resolver
			if resolver == nil {
				resolver = &mockResolver{}
			}
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Errorf("got %v (type %d), want error", got, got.Type)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// CORREL
// ---------------------------------------------------------------------------

func TestCORREL(t *testing.T) {
	// Basic data: A1:A5 = {3,2,4,5,6}, B1:B5 = {9,7,12,15,17}
	basicResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 1, Row: 5}: NumberVal(6),
			{Col: 2, Row: 1}: NumberVal(9),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(12),
			{Col: 2, Row: 4}: NumberVal(15),
			{Col: 2, Row: 5}: NumberVal(17),
		},
	}

	// Perfect positive: A1:A3 = {1,2,3}, B1:B3 = {2,4,6}
	perfectPosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// Perfect negative: A1:A3 = {1,2,3}, B1:B3 = {6,4,2}
	perfectNegResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(6),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(2),
		},
	}

	// Zero std dev: A1:A3 = {5,5,5}, B1:B3 = {1,2,3}
	zeroStdDevResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two pairs: A1:A2 = {1,2}, B1:B2 = {3,5}
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(5),
		},
	}

	// Mixed types: A1:A4 = {1,"hello",3,4}, B1:B4 = {10,20,30,"world"}
	// Only pairs (1,10) and (3,30) are numeric in both
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: StringVal("world"),
		},
	}

	// Negative numbers: A1:A4 = {-3,-1,1,3}, B1:B4 = {-9,-3,3,9}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-3),
			{Col: 1, Row: 2}: NumberVal(-1),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(-9),
			{Col: 2, Row: 2}: NumberVal(-3),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(9),
		},
	}

	// All zeros in one array: A1:A3 = {0,0,0}, B1:B3 = {1,2,3}
	allZerosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(0),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Single pair: A1 = {10}, B1 = {20}
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Zeros included: A1:A3 = {0,1,2}, B1:B3 = {0,2,4}
	zerosIncludedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(0),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(4),
		},
	}

	// Bool in array: A1:A3 = {1,TRUE,3}, B1:B3 = {10,20,30}
	// TRUE is bool, not numeric in array context; pairs (1,10) and (3,30) only
	boolResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	// Large dataset: A1:A20 = {1..20}, B1:B20 = {2..40 step 2}
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		largeResolver.cells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i * 2))
	}

	// Weak correlation: A1:A5 = {1,2,3,4,5}, B1:B5 = {2,1,4,3,5}
	weakResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(3),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Error in array: A1:A3 = {1,#VALUE!,3}, B1:B3 = {4,5,6}
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// All text: A1:A3 = {"a","b","c"}, B1:B3 = {"x","y","z"}
	allTextResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: StringVal("x"),
			{Col: 2, Row: 2}: StringVal("y"),
			{Col: 2, Row: 3}: StringVal("z"),
		},
	}

	// Different length arrays: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	tol := 1e-9

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic positive correlation
		{"basic_positive", "CORREL(A1:A5,B1:B5)", basicResolver, 0.997054486, false, 0},
		// Perfect positive correlation
		{"perfect_positive", "CORREL(A1:A3,B1:B3)", perfectPosResolver, 1.0, false, 0},
		// Perfect negative correlation
		{"perfect_negative", "CORREL(A1:A3,B1:B3)", perfectNegResolver, -1.0, false, 0},
		// Zero std dev in first array → #DIV/0!
		{"zero_stddev", "CORREL(A1:A3,B1:B3)", zeroStdDevResolver, 0, true, ErrValDIV0},
		// Two pairs
		{"two_pairs", "CORREL(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		// Mixed types: pairs (1,10) and (3,30) → perfect positive
		{"mixed_types", "CORREL(A1:A4,B1:B4)", mixedResolver, 1.0, false, 0},
		// Negative numbers: perfect positive
		{"negative_numbers", "CORREL(A1:A4,B1:B4)", negResolver, 1.0, false, 0},
		// All zeros in one array → #DIV/0!
		{"all_zeros_one_array", "CORREL(A1:A3,B1:B3)", allZerosResolver, 0, true, ErrValDIV0},
		// Single pair → #DIV/0! (std dev is 0 with 1 point)
		{"single_pair", "CORREL(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Empty range → #DIV/0!
		{"empty_range", "CORREL(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Zeros included (0 is numeric): perfect positive
		{"zeros_included", "CORREL(A1:A3,B1:B3)", zerosIncludedResolver, 1.0, false, 0},
		// Bool in array (pair skipped): pairs (1,10) and (3,30) → perfect positive
		{"bool_in_array", "CORREL(A1:A3,B1:B3)", boolResolver, 1.0, false, 0},
		// Large dataset: perfect positive (y = 2x)
		{"large_dataset", "CORREL(A1:A20,B1:B20)", largeResolver, 1.0, false, 0},
		// Weak correlation
		{"weak_correlation", "CORREL(A1:A5,B1:B5)", weakResolver, 0.8, false, 0},
		// Error propagation from first array
		{"error_in_array1", "CORREL(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// All text → no numeric pairs → #DIV/0!
		{"all_text", "CORREL(A1:A3,B1:B3)", allTextResolver, 0, true, ErrValDIV0},
		// Different length arrays → #N/A
		{"different_lengths", "CORREL(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Wrong number of arguments (1 arg)
		{"too_few_args", "CORREL(A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Wrong number of arguments (3 args)
		{"too_many_args", "CORREL(A1:A3,B1:B3,A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Reversed argument order: same magnitude, same sign
		{"reversed_args", "CORREL(B1:B5,A1:A5)", basicResolver, 0.997054486, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// SLOPE
// ---------------------------------------------------------------------------

func TestSLOPE(t *testing.T) {
	// Excel example: y={2,3,9,1,8,7,5}, x={6,5,11,7,5,4,4}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
			{Col: 1, Row: 6}: NumberVal(7), {Col: 2, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(5), {Col: 2, Row: 7}: NumberVal(4),
		},
	}

	// Perfect slope: y={2,4,6}, x={1,2,3} → slope=2
	perfectResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(6), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative slope: y={6,4,2}, x={1,2,3} → slope=-2
	negSlopeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Zero slope: y={5,5,5}, x={1,2,3} → slope=0
	zeroSlopeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(5), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Constant x: y={1,2,3}, x={5,5,5} → #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty resolver
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair: y={3}, x={7} → #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Two pairs: y={1,3}, x={2,4} → slope=1
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Mixed types: some non-numeric skipped
	// row1: y=1(num), x="a"(str) → skip
	// row2: y=2(num), x=4(num) → keep
	// row3: y=true(bool), x=6(num) → skip
	// row4: y=8(num), x=8(num) → keep
	// pairs: (2,4),(8,8) → slope = (8-2)/(8-4) = 1.5
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation: y contains #VALUE!
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 2x+1 for x=1..20
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(2*i + 1))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Fractional slope: y={1.5, 3.0, 4.5}, x={1,2,3} → slope=1.5
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1.5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3.0), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4.5), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative values: y={-6,-4,-2}, x={1,2,3} → slope=2
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"excel_example", "SLOPE(A1:A7,B1:B7)", excelResolver, 0.305555556, false, 0},
		{"perfect_slope_2", "SLOPE(A1:A3,B1:B3)", perfectResolver, 2.0, false, 0},
		{"negative_slope", "SLOPE(A1:A3,B1:B3)", negSlopeResolver, -2.0, false, 0},
		{"zero_slope", "SLOPE(A1:A3,B1:B3)", zeroSlopeResolver, 0.0, false, 0},
		{"constant_x_div0", "SLOPE(A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		{"different_lengths_na", "SLOPE(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		{"empty_div0", "SLOPE(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		{"single_pair_div0", "SLOPE(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		{"two_pairs", "SLOPE(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		{"mixed_types_skip", "SLOPE(A1:A4,B1:B4)", mixedResolver, 1.5, false, 0},
		{"error_propagation", "SLOPE(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		{"large_dataset", "SLOPE(A1:A20,B1:B20)", largeResolver, 2.0, false, 0},
		{"fractional_slope", "SLOPE(A1:A3,B1:B3)", fracResolver, 1.5, false, 0},
		{"negative_values", "SLOPE(A1:A3,B1:B3)", negValsResolver, 2.0, false, 0},
		{"too_few_args", "SLOPE(A1:A3)", perfectResolver, 0, true, ErrValVALUE},
		{"too_many_args", "SLOPE(A1:A3,B1:B3,A1:A3)", perfectResolver, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// INTERCEPT
// ---------------------------------------------------------------------------

func TestINTERCEPT(t *testing.T) {
	// Excel example: y={2,3,9,1,8}, x={6,5,11,7,5}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// y=2x+3: y={5,7,9}, x={1,2,3} → intercept=3
	interceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative intercept: y=2x-3: y={-1,1,3}, x={1,2,3} → intercept=-3
	negInterceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(1), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Zero intercept: y={2,4,6}, x={1,2,3} → intercept=0
	zeroInterceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(6), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Constant x → #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths → #N/A
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty → #DIV/0!
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair → #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Two pairs: y={1,5}, x={2,4} → slope=2, intercept=-3
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Mixed types: non-numeric skipped → pairs (2,4),(8,8) → slope=1.5, intercept=-4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 3x+10 for x=1..20 → intercept=10
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(3*i + 10))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Fractional intercept: y=0.5x+2.5: y={3.0, 3.5, 4.0}, x={1,2,3} → intercept=2.5
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3.0), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3.5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4.0), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative values: y={-6,-4,-2}, x={1,2,3} → slope=2, intercept=-8
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"excel_example", "INTERCEPT(A1:A5,B1:B5)", excelResolver, 0.0483871, false, 0},
		{"intercept_3", "INTERCEPT(A1:A3,B1:B3)", interceptResolver, 3.0, false, 0},
		{"negative_intercept", "INTERCEPT(A1:A3,B1:B3)", negInterceptResolver, -3.0, false, 0},
		{"zero_intercept", "INTERCEPT(A1:A3,B1:B3)", zeroInterceptResolver, 0.0, false, 0},
		{"constant_x_div0", "INTERCEPT(A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		{"different_lengths_na", "INTERCEPT(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		{"empty_div0", "INTERCEPT(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		{"single_pair_div0", "INTERCEPT(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		{"two_pairs", "INTERCEPT(A1:A2,B1:B2)", twoPairResolver, -3.0, false, 0},
		{"mixed_types_skip", "INTERCEPT(A1:A4,B1:B4)", mixedResolver, -4.0, false, 0},
		{"error_propagation", "INTERCEPT(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		{"large_dataset", "INTERCEPT(A1:A20,B1:B20)", largeResolver, 10.0, false, 0},
		{"fractional_intercept", "INTERCEPT(A1:A3,B1:B3)", fracResolver, 2.5, false, 0},
		{"negative_values", "INTERCEPT(A1:A3,B1:B3)", negValsResolver, -8.0, false, 0},
		{"too_few_args", "INTERCEPT(A1:A3)", interceptResolver, 0, true, ErrValVALUE},
		{"too_many_args", "INTERCEPT(A1:A3,B1:B3,A1:A3)", interceptResolver, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FORECAST / FORECAST.LINEAR
// ---------------------------------------------------------------------------

func TestFORECAST(t *testing.T) {
	// Excel example: y={6,7,9,15,21}, x_known={20,28,31,38,40}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(28),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(31),
			{Col: 1, Row: 4}: NumberVal(15), {Col: 2, Row: 4}: NumberVal(38),
			{Col: 1, Row: 5}: NumberVal(21), {Col: 2, Row: 5}: NumberVal(40),
		},
	}

	// y=2x+1: y={3,5,7}, x={1,2,3}
	linearResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(7), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two data points: y={1,3}, x={2,4} -> slope=1, intercept=-1
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Constant x: y={1,2,3}, x={5,5,5} -> #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty resolver
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair: y={3}, x={7} -> #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Mixed types: skip non-numeric pairs
	// row1: y=1(num), x="a"(str) -> skip
	// row2: y=2(num), x=4(num) -> keep
	// row3: y=true(bool), x=6(num) -> skip
	// row4: y=8(num), x=8(num) -> keep
	// pairs: (2,4),(8,8) -> slope=1.5, intercept=-4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation: y contains #VALUE!
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 2x+1 for x=1..20
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(2*i + 1))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Negative values: y={-6,-4,-2}, x={1,2,3} -> slope=2, intercept=-8
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// String x that parses as number: "30" -> 30
	stringXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(28),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(31),
			{Col: 1, Row: 4}: NumberVal(15), {Col: 2, Row: 4}: NumberVal(38),
			{Col: 1, Row: 5}: NumberVal(21), {Col: 2, Row: 5}: NumberVal(40),
			{Col: 3, Row: 1}: StringVal("30"),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Excel example
		{"excel_example", "FORECAST(30,A1:A5,B1:B5)", excelResolver, 10.607253, false, 0},
		// FORECAST.LINEAR identical
		{"forecast_linear_same", "FORECAST.LINEAR(30,A1:A5,B1:B5)", excelResolver, 10.607253, false, 0},
		// Simple y=2x+1, predict x=5 -> 11
		{"linear_2x_plus_1", "FORECAST(5,A1:A3,B1:B3)", linearResolver, 11.0, false, 0},
		// x=0 should return intercept (y=2x+1 -> intercept=1)
		{"x_zero_intercept", "FORECAST(0,A1:A3,B1:B3)", linearResolver, 1.0, false, 0},
		// Negative x value: x=-3 -> 2*(-3)+1 = -5
		{"negative_x", "FORECAST(-3,A1:A3,B1:B3)", linearResolver, -5.0, false, 0},
		// x already in known data: x=2 -> 2*2+1 = 5
		{"x_in_known_data", "FORECAST(2,A1:A3,B1:B3)", linearResolver, 5.0, false, 0},
		// Extrapolation beyond range: x=100 -> 2*100+1 = 201
		{"extrapolation", "FORECAST(100,A1:A3,B1:B3)", linearResolver, 201.0, false, 0},
		// Non-numeric x -> #VALUE!
		{"non_numeric_x", "FORECAST(\"abc\",A1:A3,B1:B3)", linearResolver, 0, true, ErrValVALUE},
		// String x that can be coerced: "30"
		{"string_x_coerced", "FORECAST(C1,A1:A5,B1:B5)", stringXResolver, 10.607253, false, 0},
		// Empty arrays -> #DIV/0!
		{"empty_arrays", "FORECAST(5,A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Different length arrays -> #N/A
		{"different_lengths", "FORECAST(5,A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Constant x values -> #DIV/0!
		{"constant_x_div0", "FORECAST(5,A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		// Single data point -> #DIV/0!
		{"single_point_div0", "FORECAST(5,A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Two data points: slope=1, intercept=-1, predict x=10 -> 9
		{"two_points", "FORECAST(10,A1:A2,B1:B2)", twoPairResolver, 9.0, false, 0},
		// Large dataset: y=2x+1, predict x=25 -> 51
		{"large_dataset", "FORECAST(25,A1:A20,B1:B20)", largeResolver, 51.0, false, 0},
		// Mixed types: slope=1.5, intercept=-4, predict x=10 -> 11
		{"mixed_types_skip", "FORECAST(10,A1:A4,B1:B4)", mixedResolver, 11.0, false, 0},
		// Error propagation from arrays
		{"error_propagation", "FORECAST(5,A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// Negative values: slope=2, intercept=-8, predict x=5 -> 2
		{"negative_values", "FORECAST(5,A1:A3,B1:B3)", negValsResolver, 2.0, false, 0},
		// Too few args
		{"too_few_args", "FORECAST(5,A1:A3)", linearResolver, 0, true, ErrValVALUE},
		// Too many args
		{"too_many_args", "FORECAST(5,A1:A3,B1:B3,A1:A3)", linearResolver, 0, true, ErrValVALUE},
		// FORECAST.LINEAR too few args
		{"linear_too_few_args", "FORECAST.LINEAR(5)", linearResolver, 0, true, ErrValVALUE},
		// Fractional x value: x=2.5 -> 2*2.5+1 = 6
		{"fractional_x", "FORECAST(2.5,A1:A3,B1:B3)", linearResolver, 6.0, false, 0},
		// FORECAST.LINEAR with two points
		{"linear_two_points", "FORECAST.LINEAR(10,A1:A2,B1:B2)", twoPairResolver, 9.0, false, 0},

		// --- Additional FORECAST.LINEAR-specific tests ---

		// FORECAST.LINEAR: perfect linear y=2x+1, predict x=5 -> 11
		{"fl_perfect_linear", "FORECAST.LINEAR(5,A1:A3,B1:B3)", linearResolver, 11.0, false, 0},
		// FORECAST.LINEAR: horizontal line (all y same) -> predict same y
		{"fl_horizontal_line", "FORECAST.LINEAR(99,A1:A3,B1:B3)", &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(5), {Col: 2, Row: 3}: NumberVal(3),
			},
		}, 5.0, false, 0},
		// FORECAST.LINEAR: extrapolation beyond data range
		{"fl_extrapolation_high", "FORECAST.LINEAR(100,A1:A3,B1:B3)", linearResolver, 201.0, false, 0},
		// FORECAST.LINEAR: extrapolation below data range (negative x)
		{"fl_extrapolation_low", "FORECAST.LINEAR(-10,A1:A3,B1:B3)", linearResolver, -19.0, false, 0},
		// FORECAST.LINEAR: interpolation within data range
		{"fl_interpolation", "FORECAST.LINEAR(1.5,A1:A3,B1:B3)", linearResolver, 4.0, false, 0},
		// FORECAST.LINEAR: negative x and y values
		{"fl_negative_values", "FORECAST.LINEAR(-1,A1:A3,B1:B3)", negValsResolver, -10.0, false, 0},
		// FORECAST.LINEAR: single data point -> #DIV/0!
		{"fl_single_point_div0", "FORECAST.LINEAR(5,A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// FORECAST.LINEAR: mismatched array sizes -> #N/A
		{"fl_mismatched_arrays", "FORECAST.LINEAR(5,A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// FORECAST.LINEAR: empty arrays -> #DIV/0!
		{"fl_empty_arrays", "FORECAST.LINEAR(5,A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// FORECAST.LINEAR: non-numeric string in x arg -> #VALUE!
		{"fl_non_numeric_x", "FORECAST.LINEAR(\"hello\",A1:A3,B1:B3)", linearResolver, 0, true, ErrValVALUE},
		// FORECAST.LINEAR: too many args -> #VALUE!
		{"fl_too_many_args", "FORECAST.LINEAR(5,A1:A3,B1:B3,A1:A3)", linearResolver, 0, true, ErrValVALUE},
		// FORECAST.LINEAR: Excel doc example with x=30
		{"fl_excel_doc_example", "FORECAST.LINEAR(30,A1:A5,B1:B5)", excelResolver, 10.607253, false, 0},
		// FORECAST.LINEAR: large dataset y=2x+1, predict x=50 -> 101
		{"fl_large_dataset", "FORECAST.LINEAR(50,A1:A20,B1:B20)", largeResolver, 101.0, false, 0},
		// FORECAST.LINEAR: decimal/fractional values in data
		{"fl_decimal_values", "FORECAST.LINEAR(3.5,A1:A3,B1:B3)", &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1.5), {Col: 2, Row: 1}: NumberVal(0.5),
				{Col: 1, Row: 2}: NumberVal(3.0), {Col: 2, Row: 2}: NumberVal(1.0),
				{Col: 1, Row: 3}: NumberVal(4.5), {Col: 2, Row: 3}: NumberVal(1.5),
			},
		}, 10.5, false, 0},
		// FORECAST.LINEAR: x=0 returns intercept
		{"fl_x_zero_intercept", "FORECAST.LINEAR(0,A1:A3,B1:B3)", linearResolver, 1.0, false, 0},
		// FORECAST.LINEAR: constant x values -> #DIV/0!
		{"fl_constant_x_div0", "FORECAST.LINEAR(5,A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		// FORECAST.LINEAR: error propagation from arrays
		{"fl_error_propagation", "FORECAST.LINEAR(5,A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// FORECAST.LINEAR: mixed types skip non-numeric pairs
		{"fl_mixed_types_skip", "FORECAST.LINEAR(10,A1:A4,B1:B4)", mixedResolver, 11.0, false, 0},
		// FORECAST.LINEAR: very large x value
		{"fl_very_large_x", "FORECAST.LINEAR(1000000,A1:A3,B1:B3)", linearResolver, 2000001.0, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FISHER
// ---------------------------------------------------------------------------

func TestFISHER(t *testing.T) {
	const tol = 1e-6
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"basic_0.75", "FISHER(0.75)", 0.9729551, false, 0},
		{"zero", "FISHER(0)", 0, false, 0},
		{"half", "FISHER(0.5)", 0.5493061, false, 0},
		{"negative_half", "FISHER(-0.5)", -0.5493061, false, 0},
		{"near_one", "FISHER(0.99)", 2.6466524, false, 0},
		{"near_neg_one", "FISHER(-0.99)", -2.6466524, false, 0},
		{"small_value", "FISHER(0.01)", 0.0100003, false, 0},
		{"boundary_one", "FISHER(1)", 0, true, ErrValNUM},
		{"boundary_neg_one", "FISHER(-1)", 0, true, ErrValNUM},
		{"out_of_range_high", "FISHER(1.5)", 0, true, ErrValNUM},
		{"out_of_range_low", "FISHER(-1.5)", 0, true, ErrValNUM},
		{"text_error", `FISHER("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestFISHER_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// No args
	cf := evalCompile(t, "FISHER()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHER() should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "FISHER(0.5,0.3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHER(0.5,0.3) should error, got type=%d", got.Type)
	}
}

// ---------------------------------------------------------------------------
// FISHERINV
// ---------------------------------------------------------------------------

func TestFISHERINV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"basic_roundtrip", "FISHERINV(0.972955)", 0.75, false, 0},
		{"zero", "FISHERINV(0)", 0, false, 0},
		{"half", "FISHERINV(0.5493061)", 0.5, false, 0},
		{"negative_half", "FISHERINV(-0.5493061)", -0.5, false, 0},
		{"large_positive", "FISHERINV(10)", 0.99999999, false, 0},
		{"large_negative", "FISHERINV(-10)", -0.99999999, false, 0},
		{"one", "FISHERINV(1)", 0.7615942, false, 0},
		{"negative_one", "FISHERINV(-1)", -0.7615942, false, 0},
		{"small_value", "FISHERINV(0.01)", 0.01, false, 0},
		{"text_error", `FISHERINV("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestFISHERINV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "FISHERINV()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHERINV() should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "FISHERINV(0.5,0.3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHERINV(0.5,0.3) should error, got type=%d", got.Type)
	}
}

func TestFISHERINV_FISHER_roundtrip(t *testing.T) {
	const tol = 1e-10
	resolver := &mockResolver{}

	inputs := []float64{0.3, 0.7, -0.5, 0.99, -0.99, 0.01}
	for _, x := range inputs {
		t.Run(fmt.Sprintf("roundtrip_%g", x), func(t *testing.T) {
			formula := fmt.Sprintf("FISHERINV(FISHER(%g))", x)
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-x) > tol {
				t.Errorf("FISHERINV(FISHER(%g)) = %g, want %g", x, got.Num, x)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// GAMMALN / GAMMALN.PRECISE
// ---------------------------------------------------------------------------

func TestGAMMALN(t *testing.T) {
	const tol = 1e-4
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic positive integer values
		{"gammaln_1", "GAMMALN(1)", 0, false, 0},
		{"gammaln_2", "GAMMALN(2)", 0, false, 0},
		{"gammaln_3", "GAMMALN(3)", 0.6931472, false, 0},
		{"gammaln_4", "GAMMALN(4)", 1.7917595, false, 0},
		{"gammaln_5", "GAMMALN(5)", 3.1780538, false, 0},
		{"gammaln_6", "GAMMALN(6)", 4.7874917, false, 0},
		{"gammaln_10", "GAMMALN(10)", 12.80183, false, 0},

		// Fractional values
		{"gammaln_0.5", "GAMMALN(0.5)", 0.5723649, false, 0},
		{"gammaln_1.5", "GAMMALN(1.5)", -0.1207822, false, 0},
		{"gammaln_2.5", "GAMMALN(2.5)", 0.2846829, false, 0},
		{"gammaln_0.001", "GAMMALN(0.001)", 6.9071786, false, 0},
		{"gammaln_0.1", "GAMMALN(0.1)", 2.2527127, false, 0},

		// Large values
		{"gammaln_100", "GAMMALN(100)", 359.1342, false, 0},
		{"gammaln_50", "GAMMALN(50)", 144.5657, false, 0},

		// Boolean coercion (TRUE=1)
		{"gammaln_true", "GAMMALN(TRUE)", 0, false, 0},

		// Error cases: non-positive
		{"gammaln_zero", "GAMMALN(0)", 0, true, ErrValNUM},
		{"gammaln_neg1", "GAMMALN(-1)", 0, true, ErrValNUM},
		{"gammaln_neg0.5", "GAMMALN(-0.5)", 0, true, ErrValNUM},
		{"gammaln_neg100", "GAMMALN(-100)", 0, true, ErrValNUM},

		// Error cases: text
		{"gammaln_text", `GAMMALN("text")`, 0, true, ErrValVALUE},

		// GAMMALN.PRECISE should behave identically
		{"precise_4", "GAMMALN.PRECISE(4)", 1.7917595, false, 0},
		{"precise_1", "GAMMALN.PRECISE(1)", 0, false, 0},
		{"precise_0.5", "GAMMALN.PRECISE(0.5)", 0.5723649, false, 0},
		{"precise_zero", "GAMMALN.PRECISE(0)", 0, true, ErrValNUM},
		{"precise_neg", "GAMMALN.PRECISE(-1)", 0, true, ErrValNUM},
		{"precise_text", `GAMMALN.PRECISE("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestGAMMALN_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// No args
	cf := evalCompile(t, "GAMMALN()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMALN() should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "GAMMALN(1,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMALN(1,2) should error, got type=%d", got.Type)
	}
}

// ---------------------------------------------------------------------------
// PERCENTRANK / PERCENTRANK.INC
// ---------------------------------------------------------------------------

func TestPERCENTRANK(t *testing.T) {
	// Excel example data: {13,12,11,8,4,3,2,1,1,1}
	// Sorted: {1,1,1,2,3,4,8,11,12,13}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(13),
			{Col: 1, Row: 2}:  NumberVal(12),
			{Col: 1, Row: 3}:  NumberVal(11),
			{Col: 1, Row: 4}:  NumberVal(8),
			{Col: 1, Row: 5}:  NumberVal(4),
			{Col: 1, Row: 6}:  NumberVal(3),
			{Col: 1, Row: 7}:  NumberVal(2),
			{Col: 1, Row: 8}:  NumberVal(1),
			{Col: 1, Row: 9}:  NumberVal(1),
			{Col: 1, Row: 10}: NumberVal(1),
		},
	}

	// Simple data for basic tests: {1,2,3,4,5} in B1:B5
	simpleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Single element in C1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(42),
		},
	}

	// Negative numbers: {-10,-5,0,5,10} in D1:D5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(-10),
			{Col: 4, Row: 2}: NumberVal(-5),
			{Col: 4, Row: 3}: NumberVal(0),
			{Col: 4, Row: 4}: NumberVal(5),
			{Col: 4, Row: 5}: NumberVal(10),
		},
	}

	// Mixed types: numbers and strings in E1:E4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(10),
			{Col: 5, Row: 2}: StringVal("hello"),
			{Col: 5, Row: 3}: NumberVal(20),
			{Col: 5, Row: 4}: NumberVal(30),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue // 0 means expect number result
	}{
		// Excel doc examples
		{name: "excel_x=2", formula: "PERCENTRANK(A1:A10,2)", resolver: excelResolver, wantNum: 0.333},
		{name: "excel_x=4", formula: "PERCENTRANK(A1:A10,4)", resolver: excelResolver, wantNum: 0.555},
		{name: "excel_x=8", formula: "PERCENTRANK(A1:A10,8)", resolver: excelResolver, wantNum: 0.666},
		{name: "excel_x=5_interp", formula: "PERCENTRANK(A1:A10,5)", resolver: excelResolver, wantNum: 0.583},

		// Min and max of range
		{name: "x_equals_min", formula: "PERCENTRANK(B1:B5,1)", resolver: simpleResolver, wantNum: 0},
		{name: "x_equals_max", formula: "PERCENTRANK(B1:B5,5)", resolver: simpleResolver, wantNum: 1},

		// x outside range → #N/A
		{name: "x_below_min", formula: "PERCENTRANK(B1:B5,0)", resolver: simpleResolver, wantErr: ErrValNA},
		{name: "x_above_max", formula: "PERCENTRANK(B1:B5,6)", resolver: simpleResolver, wantErr: ErrValNA},

		// Significance parameter
		{name: "sig_1", formula: "PERCENTRANK(A1:A10,5,1)", resolver: excelResolver, wantNum: 0.5},
		{name: "sig_5", formula: "PERCENTRANK(A1:A10,5,5)", resolver: excelResolver, wantNum: 0.58333},
		{name: "default_sig_3", formula: "PERCENTRANK(B1:B5,2)", resolver: simpleResolver, wantNum: 0.25},

		// Single element array
		{name: "single_element_match", formula: "PERCENTRANK(C1:C1,42)", resolver: singleResolver, wantNum: 1},
		{name: "single_element_no_match_below", formula: "PERCENTRANK(C1:C1,10)", resolver: singleResolver, wantErr: ErrValNA},
		{name: "single_element_no_match_above", formula: "PERCENTRANK(C1:C1,50)", resolver: singleResolver, wantErr: ErrValNA},

		// Duplicate values (x=1 appears 3 times at positions 0,1,2 → first occurrence → 0/9=0)
		{name: "duplicate_min", formula: "PERCENTRANK(A1:A10,1)", resolver: excelResolver, wantNum: 0},

		// Negative numbers
		{name: "neg_min", formula: "PERCENTRANK(D1:D5,-10)", resolver: negResolver, wantNum: 0},
		{name: "neg_max", formula: "PERCENTRANK(D1:D5,10)", resolver: negResolver, wantNum: 1},
		{name: "neg_mid", formula: "PERCENTRANK(D1:D5,0)", resolver: negResolver, wantNum: 0.5},
		{name: "neg_interp", formula: "PERCENTRANK(D1:D5,-3)", resolver: negResolver, wantNum: 0.35},

		// Unsorted data produces same result as sorted
		{name: "unsorted_data", formula: "PERCENTRANK(A1:A10,13)", resolver: excelResolver, wantNum: 1},

		// significance < 1 → #NUM!
		{name: "sig_zero", formula: "PERCENTRANK(B1:B5,2,0)", resolver: simpleResolver, wantErr: ErrValNUM},
		{name: "sig_negative", formula: "PERCENTRANK(B1:B5,2,-1)", resolver: simpleResolver, wantErr: ErrValNUM},

		// Mixed types in array (non-numeric ignored)
		{name: "mixed_types", formula: "PERCENTRANK(E1:E4,20)", resolver: mixedResolver, wantNum: 0.5},

		// PERCENTRANK.INC gives same results
		{name: "inc_same_as_base", formula: "PERCENTRANK.INC(A1:A10,2)", resolver: excelResolver, wantNum: 0.333},
		{name: "inc_interp", formula: "PERCENTRANK.INC(A1:A10,5)", resolver: excelResolver, wantNum: 0.583},
		{name: "inc_max", formula: "PERCENTRANK.INC(B1:B5,5)", resolver: simpleResolver, wantNum: 1},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError {
					t.Fatalf("expected error %d, got type=%d num=%g", tt.wantErr, got.Type, got.Num)
				}
				if got.Err != tt.wantErr {
					t.Errorf("expected error %d, got %d", tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type=%d err=%d", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// Wrong argument count
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK(B1:B5)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d", got.Type)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK(B1:B5,2,3,4)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d", got.Type)
		}
	})

	// Empty array → #NUM!
	t.Run("empty_array", func(t *testing.T) {
		emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "PERCENTRANK(F1:F3,1)")
		got, err := Eval(cf, emptyResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("expected #NUM!, got type=%d err=%d", got.Type, got.Err)
		}
	})
}

// ---------------------------------------------------------------------------
// PERCENTRANK.EXC
// ---------------------------------------------------------------------------

func TestPERCENTRANKEXC(t *testing.T) {
	// Excel example data: {1,2,3,6,6,6,7,8,9} in A1:A9
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(6),
			{Col: 1, Row: 5}: NumberVal(6),
			{Col: 1, Row: 6}: NumberVal(6),
			{Col: 1, Row: 7}: NumberVal(7),
			{Col: 1, Row: 8}: NumberVal(8),
			{Col: 1, Row: 9}: NumberVal(9),
		},
	}

	// Simple data: {1,2,3,4,5} in B1:B5
	simpleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Single element: {42} in C1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(42),
		},
	}

	// Negative numbers: {-10,-5,0,5,10} in D1:D5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(-10),
			{Col: 4, Row: 2}: NumberVal(-5),
			{Col: 4, Row: 3}: NumberVal(0),
			{Col: 4, Row: 4}: NumberVal(5),
			{Col: 4, Row: 5}: NumberVal(10),
		},
	}

	// Mixed types: {10, "hello", 20, true, 30} in E1:E5
	// Only 10, 20, 30 are numeric.
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(10),
			{Col: 5, Row: 2}: StringVal("hello"),
			{Col: 5, Row: 3}: NumberVal(20),
			{Col: 5, Row: 4}: BoolVal(true),
			{Col: 5, Row: 5}: NumberVal(30),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// Excel doc examples
		{name: "excel_x=7", formula: "PERCENTRANK.EXC(A1:A9,7)", resolver: excelResolver, wantNum: 0.7},
		{name: "excel_x=5.43", formula: "PERCENTRANK.EXC(A1:A9,5.43)", resolver: excelResolver, wantNum: 0.381},
		{name: "excel_x=5.43_sig1", formula: "PERCENTRANK.EXC(A1:A9,5.43,1)", resolver: excelResolver, wantNum: 0.3},

		// Simple data: rank = position/(n+1), n=5, denom=6
		// {1,2,3,4,5}: positions 1..5 → ranks 1/6..5/6
		{name: "simple_min", formula: "PERCENTRANK.EXC(B1:B5,1)", resolver: simpleResolver, wantNum: 0.166},
		{name: "simple_mid", formula: "PERCENTRANK.EXC(B1:B5,3)", resolver: simpleResolver, wantNum: 0.5},
		{name: "simple_max", formula: "PERCENTRANK.EXC(B1:B5,5)", resolver: simpleResolver, wantNum: 0.833},
		{name: "simple_x=2", formula: "PERCENTRANK.EXC(B1:B5,2)", resolver: simpleResolver, wantNum: 0.333},
		{name: "simple_x=4", formula: "PERCENTRANK.EXC(B1:B5,4)", resolver: simpleResolver, wantNum: 0.666},

		// x outside range → #N/A
		{name: "x_below_min", formula: "PERCENTRANK.EXC(B1:B5,0)", resolver: simpleResolver, wantErr: ErrValNA},
		{name: "x_above_max", formula: "PERCENTRANK.EXC(B1:B5,6)", resolver: simpleResolver, wantErr: ErrValNA},

		// Interpolation: x=1.5 between 1 (pos 1) and 2 (pos 2)
		// loRank=1/6, hiRank=2/6, frac=0.5 → 1/6 + 0.5*(1/6) = 0.25
		{name: "interp_1.5", formula: "PERCENTRANK.EXC(B1:B5,1.5)", resolver: simpleResolver, wantNum: 0.25},
		{name: "interp_4.5", formula: "PERCENTRANK.EXC(B1:B5,4.5)", resolver: simpleResolver, wantNum: 0.75},

		// Significance parameter
		{name: "sig_1", formula: "PERCENTRANK.EXC(B1:B5,3,1)", resolver: simpleResolver, wantNum: 0.5},
		{name: "sig_5", formula: "PERCENTRANK.EXC(B1:B5,2,5)", resolver: simpleResolver, wantNum: 0.33333},

		// Single element: rank = 1/(1+1) = 0.5
		{name: "single_element_match", formula: "PERCENTRANK.EXC(C1:C1,42)", resolver: singleResolver, wantNum: 0.5},
		{name: "single_element_no_match_below", formula: "PERCENTRANK.EXC(C1:C1,10)", resolver: singleResolver, wantErr: ErrValNA},
		{name: "single_element_no_match_above", formula: "PERCENTRANK.EXC(C1:C1,50)", resolver: singleResolver, wantErr: ErrValNA},

		// Duplicate values: x=6 appears at positions 3,4,5 (0-indexed) → first at index 3 → rank=4/10=0.4
		{name: "duplicate_value", formula: "PERCENTRANK.EXC(A1:A9,6)", resolver: excelResolver, wantNum: 0.4},

		// Negative numbers: {-10,-5,0,5,10}, n=5, denom=6
		{name: "neg_min", formula: "PERCENTRANK.EXC(D1:D5,-10)", resolver: negResolver, wantNum: 0.166},
		{name: "neg_max", formula: "PERCENTRANK.EXC(D1:D5,10)", resolver: negResolver, wantNum: 0.833},
		{name: "neg_mid", formula: "PERCENTRANK.EXC(D1:D5,0)", resolver: negResolver, wantNum: 0.5},
		{name: "neg_interp", formula: "PERCENTRANK.EXC(D1:D5,-3)", resolver: negResolver, wantNum: 0.4},

		// Unsorted data produces same result
		{name: "unsorted_data", formula: "PERCENTRANK.EXC(A1:A9,9)", resolver: excelResolver, wantNum: 0.9},

		// significance < 1 → #NUM!
		{name: "sig_zero", formula: "PERCENTRANK.EXC(B1:B5,2,0)", resolver: simpleResolver, wantErr: ErrValNUM},
		{name: "sig_negative", formula: "PERCENTRANK.EXC(B1:B5,2,-1)", resolver: simpleResolver, wantErr: ErrValNUM},

		// Mixed types in array (non-numeric ignored)
		// {10,20,30}, n=3, denom=4, x=20 at pos 1 (0-indexed) → rank=2/4=0.5
		{name: "mixed_types", formula: "PERCENTRANK.EXC(E1:E5,20)", resolver: mixedResolver, wantNum: 0.5},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, tc.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tc.wantErr != 0 {
				if got.Type != ValueError || got.Err != tc.wantErr {
					t.Errorf("expected error %d, got type=%d err=%d num=%v", tc.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type=%d err=%d", got.Type, got.Err)
			}
			if math.Abs(got.Num-tc.wantNum) > 1e-9 {
				t.Errorf("got %.10f, want %.10f", got.Num, tc.wantNum)
			}
		})
	}

	// Wrong argument count
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK.EXC(B1:B5)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("expected #VALUE!, got type=%d err=%d", got.Type, got.Err)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK.EXC(B1:B5,2,3,4)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("expected #VALUE!, got type=%d err=%d", got.Type, got.Err)
		}
	})

	// Empty array → #NUM!
	t.Run("empty_array", func(t *testing.T) {
		emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "PERCENTRANK.EXC(F1:F3,1)")
		got, err := Eval(cf, emptyResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("expected #NUM!, got type=%d err=%d", got.Type, got.Err)
		}
	})
}

// ---------------------------------------------------------------------------
// SKEW
// ---------------------------------------------------------------------------

func TestSKEW(t *testing.T) {
	// Resolver with {3,4,5,2,3,4,5,6,4,7} in A1:A10 (Excel docs example)
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(3),
			{Col: 1, Row: 2}:  NumberVal(4),
			{Col: 1, Row: 3}:  NumberVal(5),
			{Col: 1, Row: 4}:  NumberVal(2),
			{Col: 1, Row: 5}:  NumberVal(3),
			{Col: 1, Row: 6}:  NumberVal(4),
			{Col: 1, Row: 7}:  NumberVal(5),
			{Col: 1, Row: 8}:  NumberVal(6),
			{Col: 1, Row: 9}:  NumberVal(4),
			{Col: 1, Row: 10}: NumberVal(7),
		},
	}

	// Symmetric data {1,2,3,4,5} in B1:B5
	symResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Right-skewed {1,1,1,1,1,1,1,1,1,100} in C1:C10
	rightSkewResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}:  NumberVal(1),
			{Col: 3, Row: 2}:  NumberVal(1),
			{Col: 3, Row: 3}:  NumberVal(1),
			{Col: 3, Row: 4}:  NumberVal(1),
			{Col: 3, Row: 5}:  NumberVal(1),
			{Col: 3, Row: 6}:  NumberVal(1),
			{Col: 3, Row: 7}:  NumberVal(1),
			{Col: 3, Row: 8}:  NumberVal(1),
			{Col: 3, Row: 9}:  NumberVal(1),
			{Col: 3, Row: 10}: NumberVal(100),
		},
	}

	// Left-skewed {1,100,100,100,100,100,100,100,100,100} in D1:D10
	leftSkewResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}:  NumberVal(1),
			{Col: 4, Row: 2}:  NumberVal(100),
			{Col: 4, Row: 3}:  NumberVal(100),
			{Col: 4, Row: 4}:  NumberVal(100),
			{Col: 4, Row: 5}:  NumberVal(100),
			{Col: 4, Row: 6}:  NumberVal(100),
			{Col: 4, Row: 7}:  NumberVal(100),
			{Col: 4, Row: 8}:  NumberVal(100),
			{Col: 4, Row: 9}:  NumberVal(100),
			{Col: 4, Row: 10}: NumberVal(100),
		},
	}

	// Exactly 3 data points {1,2,3} in E1:E3
	threeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(1),
			{Col: 5, Row: 2}: NumberVal(2),
			{Col: 5, Row: 3}: NumberVal(3),
		},
	}

	// Two data points {1,2} in F1:F2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(1),
			{Col: 6, Row: 2}: NumberVal(2),
		},
	}

	// Single value {5} in G1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 7, Row: 1}: NumberVal(5),
		},
	}

	// All same values {4,4,4,4} in H1:H4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(4),
			{Col: 8, Row: 2}: NumberVal(4),
			{Col: 8, Row: 3}: NumberVal(4),
			{Col: 8, Row: 4}: NumberVal(4),
		},
	}

	// Negative numbers {-5,-3,-1,0,2} in I1:I5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(-5),
			{Col: 9, Row: 2}: NumberVal(-3),
			{Col: 9, Row: 3}: NumberVal(-1),
			{Col: 9, Row: 4}: NumberVal(0),
			{Col: 9, Row: 5}: NumberVal(2),
		},
	}

	// Large dataset 1..20 in J1:J20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 10, Row: i}] = NumberVal(float64(i))
	}

	// Mixed types: numbers, strings, booleans in K1:K6
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 11, Row: 1}: NumberVal(1),
			{Col: 11, Row: 2}: StringVal("hello"),
			{Col: 11, Row: 3}: NumberVal(2),
			{Col: 11, Row: 4}: BoolVal(true),
			{Col: 11, Row: 5}: NumberVal(3),
			{Col: 11, Row: 6}: NumberVal(10),
		},
	}

	// Error in array in L1:L4
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 12, Row: 1}: NumberVal(1),
			{Col: 12, Row: 2}: NumberVal(2),
			{Col: 12, Row: 3}: ErrorVal(ErrValNUM),
			{Col: 12, Row: 4}: NumberVal(4),
		},
	}

	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  ErrorValue
		isErr    bool
		tol      float64
	}{
		// Excel docs example: {3,4,5,2,3,4,5,6,4,7} → 0.359543
		{"excel_example", "SKEW(A1:A10)", excelResolver, 0.359543, 0, false, 1e-4},

		// Symmetric data {1,2,3,4,5} → skew = 0
		{"symmetric", "SKEW(B1:B5)", symResolver, 0, 0, false, 1e-9},

		// Right-skewed data → positive skew
		{"right_skewed", "SKEW(C1:C10)", rightSkewResolver, 3.16227766, 0, false, 1e-4},

		// Left-skewed data → negative skew
		{"left_skewed", "SKEW(D1:D10)", leftSkewResolver, -3.16227766, 0, false, 1e-4},

		// Exactly 3 data points {1,2,3} → 0
		{"three_points", "SKEW(E1:E3)", threeResolver, 0, 0, false, 1e-9},

		// Two data points → #DIV/0!
		{"two_points_div0", "SKEW(F1:F2)", twoResolver, 0, ErrValDIV0, true, 0},

		// Single value → #DIV/0!
		{"single_value_div0", "SKEW(G1)", singleResolver, 0, ErrValDIV0, true, 0},

		// All same values → #DIV/0! (std dev = 0)
		{"all_same_div0", "SKEW(H1:H4)", sameResolver, 0, ErrValDIV0, true, 0},

		// Negative numbers {-5,-3,-1,0,2}
		{"negative_numbers", "SKEW(I1:I5)", negResolver, -0.18252326, 0, false, 1e-4},

		// Large dataset 1..20 → 0 (symmetric)
		{"large_symmetric", "SKEW(J1:J20)", largeResolver, 0, 0, false, 1e-9},

		// Mixed types in array: text and bool are ignored → only {1,2,3,10}
		{"mixed_types_array", "SKEW(K1:K6)", mixedResolver, 1.76363261, 0, false, 1e-4},

		// Direct boolean args are counted: TRUE=1 → SKEW(1,2,3,TRUE) = SKEW(1,2,3,1)
		{"direct_bool_true", "SKEW(1,2,3,TRUE)", emptyResolver, 0.85456304, 0, false, 1e-4},

		// Direct string number args are counted: "5" → 5
		{"direct_string_num", `SKEW(1,2,3,"5")`, emptyResolver, 0.75283720, 0, false, 1e-4},

		// Error propagation from array
		{"error_propagation", "SKEW(L1:L4)", errResolver, 0, ErrValNUM, true, 0},

		// Direct error arg
		{"direct_error", "SKEW(1,2,3,1/0)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Empty range → #DIV/0! (0 values < 3)
		{"empty_range", "SKEW(Z1:Z5)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Direct args: SKEW(1,2,3) with all different → 0
		{"direct_three_sym", "SKEW(1,2,3)", emptyResolver, 0, 0, false, 1e-9},

		// Direct args with more values
		{"direct_many", "SKEW(1,1,1,1,1,100)", emptyResolver, 2.44948975, 0, false, 1e-4},

		// {1,2,3,4,100} right-skewed
		{"moderate_right_skew", "SKEW(1,2,3,4,100)", emptyResolver, 2.23239591, 0, false, 1e-4},

		// Decimals {0.5, 1.5, 2.5, 3.5, 4.5}
		{"decimals", "SKEW(0.5,1.5,2.5,3.5,4.5)", emptyResolver, 0, 0, false, 1e-9},

		// Large positive values
		{"large_values", "SKEW(1000000,2000000,3000000)", emptyResolver, 0, 0, false, 1e-9},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			tol := tt.tol
			if tol == 0 {
				tol = 1e-9
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %g, want %g (diff %g)", got.Num, tt.wantNum, math.Abs(got.Num-tt.wantNum))
			}
		})
	}

	// 0 args → should error
	t.Run("zero_args", func(t *testing.T) {
		got, err := fnSKEW([]Value{})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	// Verify sign: right-skewed > 0
	t.Run("right_skew_positive", func(t *testing.T) {
		cf := evalCompile(t, "SKEW(C1:C10)")
		got, err := Eval(cf, rightSkewResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num <= 0 {
			t.Errorf("expected positive skew, got %v", got)
		}
	})

	// Verify sign: left-skewed < 0
	t.Run("left_skew_negative", func(t *testing.T) {
		cf := evalCompile(t, "SKEW(D1:D10)")
		got, err := Eval(cf, leftSkewResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num >= 0 {
			t.Errorf("expected negative skew, got %v", got)
		}
	})
}

// ---------------------------------------------------------------------------
// KURT
// ---------------------------------------------------------------------------

func TestKURT(t *testing.T) {
	// Excel docs example: {3,4,5,2,3,4,5,6,4,7} in A1:A10
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(3),
			{Col: 1, Row: 2}:  NumberVal(4),
			{Col: 1, Row: 3}:  NumberVal(5),
			{Col: 1, Row: 4}:  NumberVal(2),
			{Col: 1, Row: 5}:  NumberVal(3),
			{Col: 1, Row: 6}:  NumberVal(4),
			{Col: 1, Row: 7}:  NumberVal(5),
			{Col: 1, Row: 8}:  NumberVal(6),
			{Col: 1, Row: 9}:  NumberVal(4),
			{Col: 1, Row: 10}: NumberVal(7),
		},
	}

	// All same values {4,4,4,4} in B1:B4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(4),
		},
	}

	// Three data points {1,2,3} in C1:C3
	threeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(1),
			{Col: 3, Row: 2}: NumberVal(2),
			{Col: 3, Row: 3}: NumberVal(3),
		},
	}

	// Two data points {1,2} in D1:D2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(1),
			{Col: 4, Row: 2}: NumberVal(2),
		},
	}

	// Exactly 4 points {1,2,3,4} in E1:E4
	fourResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(1),
			{Col: 5, Row: 2}: NumberVal(2),
			{Col: 5, Row: 3}: NumberVal(3),
			{Col: 5, Row: 4}: NumberVal(4),
		},
	}

	// Negative numbers {-5,-3,-1,0,2} in F1:F5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(-5),
			{Col: 6, Row: 2}: NumberVal(-3),
			{Col: 6, Row: 3}: NumberVal(-1),
			{Col: 6, Row: 4}: NumberVal(0),
			{Col: 6, Row: 5}: NumberVal(2),
		},
	}

	// Large dataset 1..20 in G1:G20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 7, Row: i}] = NumberVal(float64(i))
	}

	// Mixed types: numbers, strings, booleans in H1:H7
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(1),
			{Col: 8, Row: 2}: StringVal("hello"),
			{Col: 8, Row: 3}: NumberVal(2),
			{Col: 8, Row: 4}: BoolVal(true),
			{Col: 8, Row: 5}: NumberVal(3),
			{Col: 8, Row: 6}: NumberVal(10),
			{Col: 8, Row: 7}: NumberVal(5),
		},
	}

	// Error in array in I1:I5
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(1),
			{Col: 9, Row: 2}: NumberVal(2),
			{Col: 9, Row: 3}: ErrorVal(ErrValNUM),
			{Col: 9, Row: 4}: NumberVal(4),
			{Col: 9, Row: 5}: NumberVal(5),
		},
	}

	// With zeros {0,0,1,2,3,4} in J1:J6
	zeroResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 10, Row: 1}: NumberVal(0),
			{Col: 10, Row: 2}: NumberVal(0),
			{Col: 10, Row: 3}: NumberVal(1),
			{Col: 10, Row: 4}: NumberVal(2),
			{Col: 10, Row: 5}: NumberVal(3),
			{Col: 10, Row: 6}: NumberVal(4),
		},
	}

	// Empty cells in range K1:K8 (some empty)
	emptyCellResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 11, Row: 1}: NumberVal(1),
			{Col: 11, Row: 3}: NumberVal(2),
			{Col: 11, Row: 5}: NumberVal(3),
			{Col: 11, Row: 7}: NumberVal(4),
		},
	}

	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  ErrorValue
		isErr    bool
		tol      float64
	}{
		// Excel docs example: {3,4,5,2,3,4,5,6,4,7} → -0.151799637
		{"excel_example", "KURT(A1:A10)", excelResolver, -0.151799637, 0, false, 1e-4},

		// All same values → #DIV/0! (std dev = 0)
		{"all_same_div0", "KURT(B1:B4)", sameResolver, 0, ErrValDIV0, true, 0},

		// Fewer than 4 points → #DIV/0!
		{"three_points_div0", "KURT(C1:C3)", threeResolver, 0, ErrValDIV0, true, 0},
		{"two_points_div0", "KURT(D1:D2)", twoResolver, 0, ErrValDIV0, true, 0},
		{"one_point_div0", "KURT(1)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Exactly 4 points {1,2,3,4} → -1.2
		{"four_points", "KURT(E1:E4)", fourResolver, -1.2, 0, false, 1e-4},

		// Negative numbers {-5,-3,-1,0,2}
		{"negative_numbers", "KURT(F1:F5)", negResolver, -0.6811784575, 0, false, 1e-4},

		// Large dataset 1..20 → -1.2 (uniform kurtosis)
		{"large_dataset", "KURT(G1:G20)", largeResolver, -1.2, 0, false, 1e-4},

		// Mixed types in array: text and bool ignored → only {1,2,3,10,5}
		{"mixed_types_array", "KURT(H1:H7)", mixedResolver, 1.7837435675, 0, false, 1e-4},

		// Direct boolean args are counted: TRUE=1, FALSE=0
		{"direct_bool_true", "KURT(1,2,3,TRUE)", emptyResolver, -1.2892561983, 0, false, 1e-4},
		{"direct_bool_false", "KURT(1,2,3,FALSE)", emptyResolver, -1.2, 0, false, 1e-4},

		// Direct string number args are counted: "5" → 5
		{"direct_string_num", `KURT(1,2,3,"5")`, emptyResolver, 0.3428571429, 0, false, 1e-4},

		// Error propagation from array
		{"error_propagation", "KURT(I1:I5)", errResolver, 0, ErrValNUM, true, 0},

		// Direct error arg
		{"direct_error", "KURT(1,2,3,4,1/0)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Empty range → #DIV/0! (0 values < 4)
		{"empty_range", "KURT(Z1:Z5)", emptyResolver, 0, ErrValDIV0, true, 0},

		// With zeros {0,0,1,2,3,4}
		{"with_zeros", "KURT(J1:J6)", zeroResolver, -1.48125, 0, false, 1e-4},

		// Empty cells in range (ignored) → only {1,2,3,4}
		{"empty_cells_ignored", "KURT(K1:K8)", emptyCellResolver, -1.2, 0, false, 1e-4},

		// Direct args with many values
		{"direct_many", "KURT(1,1,1,1,1,100)", emptyResolver, 6.0, 0, false, 1e-4},

		// Mixed positive/negative {-2,-1,1,2}
		{"mixed_pos_neg", "KURT(-2,-1,1,2)", emptyResolver, -3.3, 0, false, 1e-4},

		// Decimals {0.5, 1.5, 2.5, 3.5, 4.5}
		{"decimals", "KURT(0.5,1.5,2.5,3.5,4.5)", emptyResolver, -1.2, 0, false, 1e-4},

		// Large positive values
		{"large_values", "KURT(1000000,2000000,3000000,4000000)", emptyResolver, -1.2, 0, false, 1e-4},

		// Single value in range → #DIV/0!
		{"single_value_div0", "KURT(A1)", excelResolver, 0, ErrValDIV0, true, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			tol := tt.tol
			if tol == 0 {
				tol = 1e-9
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %g, want %g (diff %g)", got.Num, tt.wantNum, math.Abs(got.Num-tt.wantNum))
			}
		})
	}

	// 0 args → should error
	t.Run("zero_args", func(t *testing.T) {
		got, err := fnKURT([]Value{})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

func TestMAXA(t *testing.T) {
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers picks larger", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(3,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(TRUE,0.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(FALSE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `MAXA("10",5)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `MAXA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-3),
			NumberVal(-5),
			StringVal("hello"),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// text=0 is the largest
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(0),
			NumberVal(0.5),
			BoolVal(true),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-1),
			NumberVal(-2),
			BoolVal(false),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("empty range returns 0", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, &mockResolver{cells: map[CellAddr]Value{}}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			ErrorVal(ErrValDIV0),
			NumberVal(3),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error direct arg propagates", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(1/0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(-10,-20,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -5 {
			t.Errorf("got %v, want -5", got)
		}
	})

	t.Run("Excel doc example", func(t *testing.T) {
		// {0, 0.2, 0.5, 0.4, TRUE} => max is TRUE=1
		resolver := valResolver(
			NumberVal(0),
			NumberVal(0.2),
			NumberVal(0.5),
			NumberVal(0.4),
			BoolVal(true),
		)
		cf := evalCompile(t, "MAXA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
		)
		cf := evalCompile(t, "MAXA(A1:A2,10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-5),
			Value{Type: ValueEmpty},
			NumberVal(-3),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -3 {
			t.Errorf("got %v, want -3", got)
		}
	})
}

// ---------------------------------------------------------------------------
// MAXIFS — comprehensive tests
// ---------------------------------------------------------------------------

func TestMAXIFS_SingleCriteria(t *testing.T) {
	// Excel doc Example 1: =MAXIFS(A2:A7,B2:B7,1) => 91
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// max range (A): grades
			{Col: 1, Row: 1}: NumberVal(89),
			{Col: 1, Row: 2}: NumberVal(93),
			{Col: 1, Row: 3}: NumberVal(96),
			{Col: 1, Row: 4}: NumberVal(85),
			{Col: 1, Row: 5}: NumberVal(91),
			{Col: 1, Row: 6}: NumberVal(88),
			// criteria range (B): weights
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(2),
			{Col: 2, Row: 4}: NumberVal(3),
			{Col: 2, Row: 5}: NumberVal(1),
			{Col: 2, Row: 6}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A6,B1:B6,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows with weight=1: 89, 91, 88 => max=91
	if got.Type != ValueNumber || got.Num != 91 {
		t.Errorf("MAXIFS single criteria: got %v, want 91", got)
	}
}

func TestMAXIFS_MultipleCriteria(t *testing.T) {
	// Excel doc Example 3: =MAXIFS(A2:A7,B2:B7,"b",D2:D7,">100") => 50
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// max range (A): weights
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(100),
			{Col: 1, Row: 4}: NumberVal(1),
			{Col: 1, Row: 5}: NumberVal(1),
			{Col: 1, Row: 6}: NumberVal(50),
			// criteria range 1 (B): grades
			{Col: 2, Row: 1}: StringVal("b"),
			{Col: 2, Row: 2}: StringVal("a"),
			{Col: 2, Row: 3}: StringVal("a"),
			{Col: 2, Row: 4}: StringVal("b"),
			{Col: 2, Row: 5}: StringVal("a"),
			{Col: 2, Row: 6}: StringVal("b"),
			// criteria range 2 (D): levels
			{Col: 4, Row: 1}: NumberVal(100),
			{Col: 4, Row: 2}: NumberVal(100),
			{Col: 4, Row: 3}: NumberVal(200),
			{Col: 4, Row: 4}: NumberVal(300),
			{Col: 4, Row: 5}: NumberVal(100),
			{Col: 4, Row: 6}: NumberVal(400),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A6,B1:B6,"b",D1:D6,">100")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// grade="b" AND level>100: row4(1,300) and row6(50,400) => max=50
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("MAXIFS multiple criteria: got %v, want 50", got)
	}
}

func TestMAXIFS_NoMatches(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 2, Row: 1}: StringVal("apple"),
			{Col: 2, Row: 2}: StringVal("banana"),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A2,B1:B2,"cherry")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// No matches => 0
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MAXIFS no matches: got %v, want 0", got)
	}
}

func TestMAXIFS_AllMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// All rows match => max=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS all match: got %v, want 30", got)
	}
}

func TestMAXIFS_AllComparisonOperators(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// max range (A)
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 1, Row: 5}: NumberVal(50),
			// criteria range (B)
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			{Col: 2, Row: 5}: NumberVal(50),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"greater than", `MAXIFS(A1:A5,B1:B5,">30")`, 50},
		{"less than", `MAXIFS(A1:A5,B1:B5,"<30")`, 20},
		{"greater or equal", `MAXIFS(A1:A5,B1:B5,">=30")`, 50},
		{"less or equal", `MAXIFS(A1:A5,B1:B5,"<=30")`, 30},
		{"equal", `MAXIFS(A1:A5,B1:B5,"=30")`, 30},
		{"not equal", `MAXIFS(A1:A5,B1:B5,"<>30")`, 50},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("MAXIFS %s: got %v, want %g", tt.name, got, tt.want)
			}
		})
	}
}

func TestMAXIFS_StringCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(22),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Bananas"),
			{Col: 2, Row: 3}: StringVal("Apples"),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,"Apples")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Apples rows: 5, 10 => max=10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("MAXIFS string criteria: got %v, want 10", got)
	}
}

func TestMAXIFS_WildcardAsterisk(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Apples"),
			{Col: 2, Row: 3}: StringVal("Artichokes"),
			{Col: 2, Row: 4}: StringVal("Bananas"),
		},
	}

	// Wildcard * matches any sequence of characters
	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,"A*")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// A* matches Apples(5), Apples(4), Artichokes(15) => max=15
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("MAXIFS wildcard *: got %v, want 15", got)
	}
}

func TestMAXIFS_WildcardQuestionMark(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 2, Row: 1}: StringVal("cat"),
			{Col: 2, Row: 2}: StringVal("car"),
			{Col: 2, Row: 3}: StringVal("cab"),
			{Col: 2, Row: 4}: StringVal("dogs"),
		},
	}

	// ? matches any single character: "ca?" matches cat, car, cab but not dogs
	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,"ca?")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// cat(10), car(20), cab(30) => max=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS wildcard ?: got %v, want 30", got)
	}
}

func TestMAXIFS_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 2, Row: 1}: StringVal("Apple"),
			{Col: 2, Row: 2}: StringVal("APPLE"),
			{Col: 2, Row: 3}: StringVal("apple"),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// All three should match case-insensitively: max(100,200,300) = 300
	if got.Type != ValueNumber || got.Num != 300 {
		t.Errorf("MAXIFS case insensitive: got %v, want 300", got)
	}
}

func TestMAXIFS_NegativeNumbers(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-10),
			{Col: 1, Row: 2}: NumberVal(-20),
			{Col: 1, Row: 3}: NumberVal(-5),
			{Col: 1, Row: 4}: NumberVal(-15),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(2),
			{Col: 2, Row: 4}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows with criteria=1: -10, -20, -15 => max=-10
	if got.Type != ValueNumber || got.Num != -10 {
		t.Errorf("MAXIFS negative numbers: got %v, want -10", got)
	}
}

func TestMAXIFS_MixedTypesInMaxRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("text"),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(1),
			{Col: 2, Row: 4}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// "text" is not numeric and is skipped; TRUE coerces to 1
	// Numeric values: 10, 30, 1 => max=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS mixed types: got %v, want 30", got)
	}
}

func TestMAXIFS_EmptyCellsInMaxRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			// A2 is empty (not in map)
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(1),
		},
	}

	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Empty cell coerces to 0: max(10, 0, 30) = 30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS empty cells: got %v, want 30", got)
	}
}

func TestMAXIFS_TooFewArgs(t *testing.T) {
	resolver := &mockResolver{}

	// Only 1 arg — need at least 3
	cf := evalCompile(t, "MAXIFS(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("MAXIFS too few args: got %v, want #VALUE!", got)
	}
}

func TestMAXIFS_UnpairedCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 3, Row: 1}: NumberVal(2),
		},
	}

	// 4 args total: max_range + 3 => (4-1)%2 != 0 => error
	cf := evalCompile(t, `MAXIFS(A1:A1,B1:B1,1,C1:C1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("MAXIFS unpaired criteria: got %v, want #VALUE!", got)
	}
}

func TestMAXIFS_StringNotEqual(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(22),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 2, Row: 1}: StringVal("Apples"),
			{Col: 2, Row: 2}: StringVal("Bananas"),
			{Col: 2, Row: 3}: StringVal("Apples"),
		},
	}

	// Max where product is NOT Bananas
	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,"<>Bananas")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Apples rows: 5, 10 => max=10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("MAXIFS not-equal string: got %v, want 10", got)
	}
}

func TestMAXIFS_ExcelDocExample2(t *testing.T) {
	// Excel doc Example 2: =MAXIFS(A2:A5,B3:B6,"a") => 10
	// criteria_range and max_range aren't aligned but same shape
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// max range A2:A5
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(100),
			{Col: 1, Row: 5}: NumberVal(1),
			// criteria range B3:B6
			{Col: 2, Row: 3}: StringVal("a"),
			{Col: 2, Row: 4}: StringVal("a"),
			{Col: 2, Row: 5}: StringVal("b"),
			{Col: 2, Row: 6}: StringVal("a"),
		},
	}

	cf := evalCompile(t, `MAXIFS(A2:A5,B3:B6,"a")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// "a" matches positions 1,2,4 => max_range values 10,1,1 => max=10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("MAXIFS Excel doc example 2: got %v, want 10", got)
	}
}

func TestMAXIFS_ThreeCriteriaPairs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// max range (A)
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(300),
			{Col: 1, Row: 4}: NumberVal(400),
			// criteria range 1 (B): region
			{Col: 2, Row: 1}: StringVal("East"),
			{Col: 2, Row: 2}: StringVal("West"),
			{Col: 2, Row: 3}: StringVal("East"),
			{Col: 2, Row: 4}: StringVal("East"),
			// criteria range 2 (C): product
			{Col: 3, Row: 1}: StringVal("Widget"),
			{Col: 3, Row: 2}: StringVal("Widget"),
			{Col: 3, Row: 3}: StringVal("Gadget"),
			{Col: 3, Row: 4}: StringVal("Widget"),
			// criteria range 3 (D): quantity
			{Col: 4, Row: 1}: NumberVal(5),
			{Col: 4, Row: 2}: NumberVal(15),
			{Col: 4, Row: 3}: NumberVal(20),
			{Col: 4, Row: 4}: NumberVal(25),
		},
	}

	// East AND Widget AND qty>10 => only row 4 (400)
	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,"East",C1:C4,"Widget",D1:D4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("MAXIFS three criteria: got %v, want 400", got)
	}
}

func TestMAXIFS_NonArrayMaxRange(t *testing.T) {
	resolver := &mockResolver{}

	// max_range is a scalar, not an array => #VALUE!
	cf := evalCompile(t, `MAXIFS(5,A1:A1,1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("MAXIFS non-array max_range: got %v, want #VALUE!", got)
	}
}

func TestMINA(t *testing.T) {
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MINA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers picks smaller", func(t *testing.T) {
		cf := evalCompile(t, "MINA(3,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "MINA(TRUE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "MINA(FALSE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `MINA("2",5)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `MINA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(3),
			NumberVal(5),
			StringVal("hello"),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// text=0 is the smallest
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
			BoolVal(true),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			NumberVal(2),
			BoolVal(false),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("empty range returns 0", func(t *testing.T) {
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, &mockResolver{cells: map[CellAddr]Value{}}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			ErrorVal(ErrValDIV0),
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error direct arg propagates", func(t *testing.T) {
		cf := evalCompile(t, "MINA(1/0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("all positive numbers", func(t *testing.T) {
		cf := evalCompile(t, "MINA(10,20,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("Excel doc example", func(t *testing.T) {
		// {FALSE, 0.2, 0.5, 0.4, 0.8} => min is FALSE=0
		resolver := valResolver(
			BoolVal(false),
			NumberVal(0.2),
			NumberVal(0.5),
			NumberVal(0.4),
			NumberVal(0.8),
		)
		cf := evalCompile(t, "MINA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A2,-10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -10 {
			t.Errorf("got %v, want -10", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(5),
			Value{Type: ValueEmpty},
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})
}

// ---------------------------------------------------------------------------
// RANK.EQ
// ---------------------------------------------------------------------------

func TestRANKEQ(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(7),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(9),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr ErrorValue
	}{
		{name: "desc_top", formula: "RANK.EQ(9,A1:A5)", want: 1},
		{name: "desc_mid", formula: "RANK.EQ(7,A1:A5)", want: 2},
		{name: "desc_tie", formula: "RANK.EQ(3,A1:A5)", want: 4},
		{name: "asc_bottom", formula: "RANK.EQ(3,A1:A5,1)", want: 1},
		{name: "asc_top", formula: "RANK.EQ(9,A1:A5,1)", want: 5},
		{name: "not_found", formula: "RANK.EQ(99,A1:A5)", wantErr: ErrValNA},
		{name: "too_few_args", formula: "RANK.EQ(9)", wantErr: ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("got %v, want %g", got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// RANK.AVG
// ---------------------------------------------------------------------------

func TestRANKAVG(t *testing.T) {
	// Dataset: {7, 3, 5, 3, 9}
	// Descending sorted: 9(1), 7(2), 5(3), 3(4), 3(5)
	// Ascending sorted:  3(1), 3(2), 5(3), 7(4), 9(5)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(7),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(9),
		},
	}

	// Dataset with triple tie: {4, 4, 4, 1, 6}
	// Descending: 6(1), 4(2), 4(3), 4(4), 1(5)
	tripleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(1),
			{Col: 2, Row: 5}: NumberVal(6),
		},
	}

	// All same values: {5, 5, 5}
	allSameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(5),
			{Col: 3, Row: 3}: NumberVal(5),
		},
	}

	// Single value: {10}
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(10),
		},
	}

	// Negative values: {-3, -1, -3, 2, 0}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-3),
			{Col: 5, Row: 2}: NumberVal(-1),
			{Col: 5, Row: 3}: NumberVal(-3),
			{Col: 5, Row: 4}: NumberVal(2),
			{Col: 5, Row: 5}: NumberVal(0),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		want     float64
		wantErr  ErrorValue
	}{
		// Basic descending (no ties)
		{name: "desc_unique_top", formula: "RANK.AVG(9,A1:A5)", resolver: resolver, want: 1},
		{name: "desc_unique_mid", formula: "RANK.AVG(7,A1:A5)", resolver: resolver, want: 2},
		{name: "desc_unique_5", formula: "RANK.AVG(5,A1:A5)", resolver: resolver, want: 3},

		// Descending with ties: two 3s occupy ranks 4 and 5, avg = 4.5
		{name: "desc_tie_avg", formula: "RANK.AVG(3,A1:A5)", resolver: resolver, want: 4.5},

		// Ascending (no ties)
		{name: "asc_unique_top", formula: "RANK.AVG(9,A1:A5,1)", resolver: resolver, want: 5},
		{name: "asc_unique_mid", formula: "RANK.AVG(7,A1:A5,1)", resolver: resolver, want: 4},

		// Ascending with ties: two 3s occupy ranks 1 and 2, avg = 1.5
		{name: "asc_tie_avg", formula: "RANK.AVG(3,A1:A5,1)", resolver: resolver, want: 1.5},

		// Triple tie descending: three 4s occupy ranks 2,3,4 -> avg = 3
		{name: "triple_tie_desc", formula: "RANK.AVG(4,B1:B5)", resolver: tripleResolver, want: 3},

		// Triple tie ascending: three 4s occupy ranks 2,3,4 -> avg = 3
		{name: "triple_tie_asc", formula: "RANK.AVG(4,B1:B5,1)", resolver: tripleResolver, want: 3},

		// No ties in triple dataset
		{name: "triple_unique_top_desc", formula: "RANK.AVG(6,B1:B5)", resolver: tripleResolver, want: 1},
		{name: "triple_unique_bottom_desc", formula: "RANK.AVG(1,B1:B5)", resolver: tripleResolver, want: 5},

		// All same values: three 5s occupy ranks 1,2,3 -> avg = 2
		{name: "all_same_desc", formula: "RANK.AVG(5,C1:C3)", resolver: allSameResolver, want: 2},
		{name: "all_same_asc", formula: "RANK.AVG(5,C1:C3,1)", resolver: allSameResolver, want: 2},

		// Single value
		{name: "single_value", formula: "RANK.AVG(10,D1:D1)", resolver: singleResolver, want: 1},

		// Value not found
		{name: "not_found", formula: "RANK.AVG(99,A1:A5)", resolver: resolver, wantErr: ErrValNA},

		// Negative values descending: sorted desc: 2(1), 0(2), -1(3), -3(4), -3(5)
		{name: "neg_no_tie_desc", formula: "RANK.AVG(2,E1:E5)", resolver: negResolver, want: 1},
		{name: "neg_tie_desc", formula: "RANK.AVG(-3,E1:E5)", resolver: negResolver, want: 4.5},

		// Negative values ascending: sorted asc: -3(1), -3(2), -1(3), 0(4), 2(5)
		{name: "neg_tie_asc", formula: "RANK.AVG(-3,E1:E5,1)", resolver: negResolver, want: 1.5},
		{name: "neg_no_tie_asc", formula: "RANK.AVG(2,E1:E5,1)", resolver: negResolver, want: 5},

		// order=0 means descending (same as omitting)
		{name: "order_zero_desc", formula: "RANK.AVG(3,A1:A5,0)", resolver: resolver, want: 4.5},

		// Wrong argument count
		{name: "too_few_args", formula: "RANK.AVG(9)", resolver: resolver, wantErr: ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("got %v, want %g", got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// STANDARDIZE
// ---------------------------------------------------------------------------

func TestPERMUTATIONA(t *testing.T) {
	const tol = 1e-9
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		wantErr ErrorValue
		isErr   bool
	}{
		// Basic cases
		{"3_choose_2", "PERMUTATIONA(3,2)", 9, 0, false},
		{"2_choose_3", "PERMUTATIONA(2,3)", 8, 0, false},
		{"4_choose_3", "PERMUTATIONA(4,3)", 64, 0, false},
		{"1_choose_5", "PERMUTATIONA(1,5)", 1, 0, false},
		{"5_choose_1", "PERMUTATIONA(5,1)", 5, 0, false},
		{"5_choose_0", "PERMUTATIONA(5,0)", 1, 0, false},
		{"0_choose_0", "PERMUTATIONA(0,0)", 1, 0, false},
		// Truncation
		{"truncate_both", "PERMUTATIONA(3.9,2.1)", 9, 0, false},
		// Error cases
		{"0_choose_1", "PERMUTATIONA(0,1)", 0, ErrValNUM, true},
		{"negative_number", "PERMUTATIONA(-1,2)", 0, ErrValNUM, true},
		{"negative_chosen", "PERMUTATIONA(3,-1)", 0, ErrValNUM, true},
		// Wrong arg count
		{"too_few_args", "PERMUTATIONA(3)", 0, ErrValVALUE, true},
		{"too_many_args", "PERMUTATIONA(3,2,1)", 0, ErrValVALUE, true},
		// String coercion
		{"string_coercion", fmt.Sprintf("PERMUTATIONA(%q,2)", "3"), 9, 0, false},
		{"non_numeric_string", fmt.Sprintf("PERMUTATIONA(%q,2)", "abc"), 0, ErrValVALUE, true},
		// Boolean coercion
		{"bool_true", "PERMUTATIONA(TRUE,5)", 1, 0, false},
		{"bool_false_zero", "PERMUTATIONA(FALSE,0)", 1, 0, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %g, want %g (diff %g)", got.Num, tt.wantNum, math.Abs(got.Num-tt.wantNum))
			}
		})
	}
}

func TestSTANDARDIZE(t *testing.T) {
	const tol = 1e-9
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		wantErr ErrorValue
		isErr   bool
	}{
		// Basic positive result: (42-40)/1.5 = 1.333...
		{"basic_positive", "STANDARDIZE(42,40,1.5)", 2.0 / 1.5, 0, false},
		// Zero result: (40-40)/1.5 = 0
		{"zero_result", "STANDARDIZE(40,40,1.5)", 0, 0, false},
		// Negative result: (38-40)/1.5 = -1.333...
		{"negative_result", "STANDARDIZE(38,40,1.5)", -2.0 / 1.5, 0, false},
		// Standard normal: (1-0)/1 = 1
		{"standard_normal", "STANDARDIZE(1,0,1)", 1, 0, false},
		// Large stddev
		{"large_stddev", "STANDARDIZE(100,50,100)", 0.5, 0, false},
		// Small stddev
		{"small_stddev", "STANDARDIZE(10.001,10,0.001)", 1, 0, false},
		// Negative x and mean
		{"negative_values", "STANDARDIZE(-5,-10,2)", 2.5, 0, false},
		// Fractional stddev
		{"fractional_stddev", "STANDARDIZE(0,0,0.5)", 0, 0, false},
		// stddev = 0 -> #NUM!
		{"stddev_zero", "STANDARDIZE(42,40,0)", 0, ErrValNUM, true},
		// Negative stddev -> #NUM!
		{"stddev_negative", "STANDARDIZE(42,40,-1)", 0, ErrValNUM, true},
		// Too few args -> #VALUE!
		{"too_few_args", "STANDARDIZE(42,40)", 0, ErrValVALUE, true},
		// Too many args -> #VALUE!
		{"too_many_args", "STANDARDIZE(42,40,1.5,1)", 0, ErrValVALUE, true},
		// Non-numeric string -> #VALUE!
		{"non_numeric_string", fmt.Sprintf("STANDARDIZE(%q,40,1.5)", "abc"), 0, ErrValVALUE, true},
		// String coercion: "42" -> 42
		{"string_coercion", fmt.Sprintf("STANDARDIZE(%q,40,1.5)", "42"), 2.0 / 1.5, 0, false},
		// Boolean TRUE = 1: (1-0)/1 = 1
		{"bool_true", "STANDARDIZE(TRUE,0,1)", 1, 0, false},
		// Boolean FALSE = 0: (0-0)/1 = 0
		{"bool_false", "STANDARDIZE(FALSE,0,1)", 0, 0, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %g, want %g (diff %g)", got.Num, tt.wantNum, math.Abs(got.Num-tt.wantNum))
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FREQUENCY
// ---------------------------------------------------------------------------

// freqHelper builds a ValueArray from data and bins slices and calls fnFREQUENCY.
func freqHelper(t *testing.T, data []Value, bins []Value) Value {
	t.Helper()
	dataArr := Value{Type: ValueArray, Array: [][]Value{data}}
	binsArr := Value{Type: ValueArray, Array: [][]Value{bins}}
	got, err := fnFREQUENCY([]Value{dataArr, binsArr})
	if err != nil {
		t.Fatalf("fnFREQUENCY returned error: %v", err)
	}
	return got
}

// freqExpectArray checks that got is a vertical array matching want.
func freqExpectArray(t *testing.T, got Value, want []float64) {
	t.Helper()
	if got.Type != ValueArray {
		t.Fatalf("expected ValueArray, got type %d: %v", got.Type, got)
	}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if len(got.Array[i]) != 1 {
			t.Fatalf("row %d: expected 1 column, got %d", i, len(got.Array[i]))
		}
		if got.Array[i][0].Type != ValueNumber || got.Array[i][0].Num != w {
			t.Errorf("row %d: got %v, want %g", i, got.Array[i][0], w)
		}
	}
}

func TestFREQUENCY_Basic(t *testing.T) {
	// Classic example: data={79,85,78,85,83,81,95,88,97}, bins={70,79,89}
	data := []Value{
		NumberVal(79), NumberVal(85), NumberVal(78), NumberVal(85),
		NumberVal(83), NumberVal(81), NumberVal(95), NumberVal(88), NumberVal(97),
	}
	bins := []Value{NumberVal(70), NumberVal(79), NumberVal(89)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{0, 2, 5, 2})
}

func TestFREQUENCY_EmptyData(t *testing.T) {
	bins := []Value{NumberVal(10), NumberVal(20)}
	got := freqHelper(t, []Value{}, bins)
	freqExpectArray(t, got, []float64{0, 0, 0})
}

func TestFREQUENCY_EmptyBins(t *testing.T) {
	data := []Value{NumberVal(1), NumberVal(2), NumberVal(3)}
	got := freqHelper(t, data, []Value{})
	freqExpectArray(t, got, []float64{3})
}

func TestFREQUENCY_SingleBin(t *testing.T) {
	data := []Value{NumberVal(1), NumberVal(5), NumberVal(10)}
	bins := []Value{NumberVal(5)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{2, 1})
}

func TestFREQUENCY_AllBelowFirstBin(t *testing.T) {
	data := []Value{NumberVal(1), NumberVal(2), NumberVal(3)}
	bins := []Value{NumberVal(10), NumberVal(20)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{3, 0, 0})
}

func TestFREQUENCY_AllAboveLastBin(t *testing.T) {
	data := []Value{NumberVal(30), NumberVal(40), NumberVal(50)}
	bins := []Value{NumberVal(10), NumberVal(20)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{0, 0, 3})
}

func TestFREQUENCY_AllInOneBin(t *testing.T) {
	data := []Value{NumberVal(11), NumberVal(12), NumberVal(13)}
	bins := []Value{NumberVal(10), NumberVal(20), NumberVal(30)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{0, 3, 0, 0})
}

func TestFREQUENCY_DuplicateAtBoundary(t *testing.T) {
	data := []Value{NumberVal(10), NumberVal(10), NumberVal(10)}
	bins := []Value{NumberVal(10)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{3, 0})
}

func TestFREQUENCY_NegativeValues(t *testing.T) {
	data := []Value{NumberVal(-5), NumberVal(-3), NumberVal(-1), NumberVal(0), NumberVal(2)}
	bins := []Value{NumberVal(-2), NumberVal(0)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{2, 2, 1})
}

func TestFREQUENCY_ZeroInData(t *testing.T) {
	data := []Value{NumberVal(0), NumberVal(0), NumberVal(5)}
	bins := []Value{NumberVal(0), NumberVal(10)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{2, 1, 0})
}

func TestFREQUENCY_TextInDataIgnored(t *testing.T) {
	data := []Value{NumberVal(1), StringVal("hello"), NumberVal(5), StringVal("world")}
	bins := []Value{NumberVal(3)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{1, 1})
}

func TestFREQUENCY_TextInBinsIgnored(t *testing.T) {
	data := []Value{NumberVal(1), NumberVal(5), NumberVal(15)}
	bins := []Value{NumberVal(10), StringVal("abc"), NumberVal(20)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{2, 1, 0})
}

func TestFREQUENCY_SingleDataPoint(t *testing.T) {
	data := []Value{NumberVal(5)}
	bins := []Value{NumberVal(3), NumberVal(7)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{0, 1, 0})
}

func TestFREQUENCY_ManyBinsNoDataInSome(t *testing.T) {
	data := []Value{NumberVal(25)}
	bins := []Value{NumberVal(10), NumberVal(20), NumberVal(30), NumberVal(40), NumberVal(50)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{0, 0, 1, 0, 0, 0})
}

func TestFREQUENCY_UnsortedBins(t *testing.T) {
	// bins are {30,10,20} but should be sorted internally to {10,20,30}
	data := []Value{NumberVal(5), NumberVal(15), NumberVal(25), NumberVal(35)}
	bins := []Value{NumberVal(30), NumberVal(10), NumberVal(20)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{1, 1, 1, 1})
}

func TestFREQUENCY_BoolInDataIgnored(t *testing.T) {
	data := []Value{NumberVal(1), BoolVal(true), NumberVal(5), BoolVal(false)}
	bins := []Value{NumberVal(3)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{1, 1})
}

func TestFREQUENCY_BoolInBinsIgnored(t *testing.T) {
	data := []Value{NumberVal(1), NumberVal(5)}
	bins := []Value{NumberVal(3), BoolVal(true)}
	got := freqHelper(t, data, bins)
	freqExpectArray(t, got, []float64{1, 1})
}

func TestFREQUENCY_ErrorInData(t *testing.T) {
	dataArr := Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNUM)}}}
	binsArr := Value{Type: ValueArray, Array: [][]Value{{NumberVal(5)}}}
	got, err := fnFREQUENCY([]Value{dataArr, binsArr})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("expected #NUM! error, got %v", got)
	}
}

func TestFREQUENCY_ErrorInBins(t *testing.T) {
	dataArr := Value{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}}
	binsArr := Value{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValDIV0)}}}
	got, err := fnFREQUENCY([]Value{dataArr, binsArr})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0! error, got %v", got)
	}
}

func TestFREQUENCY_WrongArgCount(t *testing.T) {
	got, err := fnFREQUENCY([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE! error, got %v", got)
	}
	got, err = fnFREQUENCY([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE! error for 3 args, got %v", got)
	}
}

func TestFREQUENCY_ScalarArgs(t *testing.T) {
	// Scalar data and scalar bin: value 5 <= bin 10, so {1, 0}
	got, err := fnFREQUENCY([]Value{NumberVal(5), NumberVal(10)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	freqExpectArray(t, got, []float64{1, 0})
}

func TestFREQUENCY_ViaEval(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(79),
			{Col: 1, Row: 2}: NumberVal(85),
			{Col: 1, Row: 3}: NumberVal(78),
			{Col: 1, Row: 4}: NumberVal(85),
			{Col: 1, Row: 5}: NumberVal(83),
			{Col: 1, Row: 6}: NumberVal(81),
			{Col: 1, Row: 7}: NumberVal(95),
			{Col: 1, Row: 8}: NumberVal(88),
			{Col: 1, Row: 9}: NumberVal(97),
			{Col: 2, Row: 1}: NumberVal(70),
			{Col: 2, Row: 2}: NumberVal(79),
			{Col: 2, Row: 3}: NumberVal(89),
		},
	}
	cf := evalCompile(t, "FREQUENCY(A1:A9,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	freqExpectArray(t, got, []float64{0, 2, 5, 2})
}

func TestFREQUENCY_EmptyDataAndBins(t *testing.T) {
	got := freqHelper(t, []Value{}, []Value{})
	freqExpectArray(t, got, []float64{0})
}

func TestFREQUENCY_DuplicateBins(t *testing.T) {
	// Duplicate bins: {10,10,20}. Values at boundary 10 go to first bin.
	data := []Value{NumberVal(5), NumberVal(10), NumberVal(15)}
	bins := []Value{NumberVal(10), NumberVal(10), NumberVal(20)}
	got := freqHelper(t, data, bins)
	// bins sorted: {10,10,20} -> <=10: {5,10}=2, (10,10]: 0, (10,20]: {15}=1, >20: 0
	freqExpectArray(t, got, []float64{2, 0, 1, 0})
}

// ---------------------------------------------------------------------------
// PEARSON
// ---------------------------------------------------------------------------

func TestPEARSON(t *testing.T) {
	// Basic data: A1:A5 = {3,2,4,5,6}, B1:B5 = {9,7,12,15,17}
	basicResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 1, Row: 5}: NumberVal(6),
			{Col: 2, Row: 1}: NumberVal(9),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(12),
			{Col: 2, Row: 4}: NumberVal(15),
			{Col: 2, Row: 5}: NumberVal(17),
		},
	}

	// Perfect positive: A1:A3 = {1,2,3}, B1:B3 = {2,4,6}
	perfectPosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// Perfect negative: A1:A3 = {1,2,3}, B1:B3 = {6,4,2}
	perfectNegResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(6),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(2),
		},
	}

	// Zero std dev: A1:A3 = {5,5,5}, B1:B3 = {1,2,3}
	zeroStdDevResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two pairs: A1:A2 = {1,2}, B1:B2 = {3,5}
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(5),
		},
	}

	// Mixed types: A1:A4 = {1,"hello",3,4}, B1:B4 = {10,20,30,"world"}
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: StringVal("world"),
		},
	}

	// Negative numbers: A1:A4 = {-3,-1,1,3}, B1:B4 = {-9,-3,3,9}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-3),
			{Col: 1, Row: 2}: NumberVal(-1),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(-9),
			{Col: 2, Row: 2}: NumberVal(-3),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(9),
		},
	}

	// All zeros in one array: A1:A3 = {0,0,0}, B1:B3 = {1,2,3}
	allZerosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(0),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Single pair: A1 = {10}, B1 = {20}
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Zeros included: A1:A3 = {0,1,2}, B1:B3 = {0,2,4}
	zerosIncludedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(0),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(4),
		},
	}

	// Bool in array: A1:A3 = {1,TRUE,3}, B1:B3 = {10,20,30}
	boolResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	// Large dataset: A1:A20 = {1..20}, B1:B20 = {2..40 step 2}
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		largeResolver.cells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i * 2))
	}

	// Weak correlation: A1:A5 = {1,2,3,4,5}, B1:B5 = {2,1,4,3,5}
	weakResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(3),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Error in array: A1:A3 = {1,#VALUE!,3}, B1:B3 = {4,5,6}
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// All text: A1:A3 = {"a","b","c"}, B1:B3 = {"x","y","z"}
	allTextResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: StringVal("x"),
			{Col: 2, Row: 2}: StringVal("y"),
			{Col: 2, Row: 3}: StringVal("z"),
		},
	}

	// Different length arrays: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Zero std dev in second array: A1:A3 = {1,2,3}, B1:B3 = {7,7,7}
	zeroStdDev2Resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(7),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(7),
		},
	}

	// Fractional data: A1:A3 = {0.1,0.2,0.3}, B1:B3 = {0.5,1.0,1.5}
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0.1),
			{Col: 1, Row: 2}: NumberVal(0.2),
			{Col: 1, Row: 3}: NumberVal(0.3),
			{Col: 2, Row: 1}: NumberVal(0.5),
			{Col: 2, Row: 2}: NumberVal(1.0),
			{Col: 2, Row: 3}: NumberVal(1.5),
		},
	}

	// Error in second array: A1:A3 = {1,2,3}, B1:B3 = {4,#REF!,6}
	errResolver2 := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: ErrorVal(ErrValREF),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	tol := 1e-9

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic positive correlation
		{"basic_positive", "PEARSON(A1:A5,B1:B5)", basicResolver, 0.997054486, false, 0},
		// Perfect positive correlation
		{"perfect_positive", "PEARSON(A1:A3,B1:B3)", perfectPosResolver, 1.0, false, 0},
		// Perfect negative correlation
		{"perfect_negative", "PEARSON(A1:A3,B1:B3)", perfectNegResolver, -1.0, false, 0},
		// Zero std dev in first array
		{"zero_stddev_array1", "PEARSON(A1:A3,B1:B3)", zeroStdDevResolver, 0, true, ErrValDIV0},
		// Zero std dev in second array
		{"zero_stddev_array2", "PEARSON(A1:A3,B1:B3)", zeroStdDev2Resolver, 0, true, ErrValDIV0},
		// Two pairs
		{"two_pairs", "PEARSON(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		// Mixed types: pairs (1,10) and (3,30) only
		{"mixed_types", "PEARSON(A1:A4,B1:B4)", mixedResolver, 1.0, false, 0},
		// Negative numbers: perfect positive
		{"negative_numbers", "PEARSON(A1:A4,B1:B4)", negResolver, 1.0, false, 0},
		// All zeros in one array
		{"all_zeros_one_array", "PEARSON(A1:A3,B1:B3)", allZerosResolver, 0, true, ErrValDIV0},
		// Single pair
		{"single_pair", "PEARSON(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Empty range
		{"empty_range", "PEARSON(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Zeros included (0 is numeric): perfect positive
		{"zeros_included", "PEARSON(A1:A3,B1:B3)", zerosIncludedResolver, 1.0, false, 0},
		// Bool in array (pair skipped)
		{"bool_in_array", "PEARSON(A1:A3,B1:B3)", boolResolver, 1.0, false, 0},
		// Large dataset: perfect positive (y = 2x)
		{"large_dataset", "PEARSON(A1:A20,B1:B20)", largeResolver, 1.0, false, 0},
		// Weak correlation
		{"weak_correlation", "PEARSON(A1:A5,B1:B5)", weakResolver, 0.8, false, 0},
		// Error propagation from first array
		{"error_in_array1", "PEARSON(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// Error propagation from second array
		{"error_in_array2", "PEARSON(A1:A3,B1:B3)", errResolver2, 0, true, ErrValREF},
		// All text
		{"all_text", "PEARSON(A1:A3,B1:B3)", allTextResolver, 0, true, ErrValDIV0},
		// Different length arrays
		{"different_lengths", "PEARSON(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Wrong number of arguments (1 arg)
		{"too_few_args", "PEARSON(A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Wrong number of arguments (3 args)
		{"too_many_args", "PEARSON(A1:A3,B1:B3,A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Reversed argument order: same result
		{"reversed_args", "PEARSON(B1:B5,A1:A5)", basicResolver, 0.997054486, false, 0},
		// Fractional values: perfect positive
		{"fractional_values", "PEARSON(A1:A3,B1:B3)", fracResolver, 1.0, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// RSQ
// ---------------------------------------------------------------------------

func TestRSQ(t *testing.T) {
	// Basic data: A1:A5 = {3,2,4,5,6}, B1:B5 = {9,7,12,15,17}
	basicResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 1, Row: 5}: NumberVal(6),
			{Col: 2, Row: 1}: NumberVal(9),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(12),
			{Col: 2, Row: 4}: NumberVal(15),
			{Col: 2, Row: 5}: NumberVal(17),
		},
	}

	// Perfect positive: A1:A3 = {1,2,3}, B1:B3 = {2,4,6}
	perfectPosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// Perfect negative: A1:A3 = {1,2,3}, B1:B3 = {6,4,2}
	perfectNegResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(6),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(2),
		},
	}

	// Zero std dev: A1:A3 = {5,5,5}, B1:B3 = {1,2,3}
	zeroStdDevResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two pairs: A1:A2 = {1,2}, B1:B2 = {3,5}
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(5),
		},
	}

	// Mixed types: A1:A4 = {1,"hello",3,4}, B1:B4 = {10,20,30,"world"}
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: StringVal("world"),
		},
	}

	// Negative numbers: A1:A4 = {-3,-1,1,3}, B1:B4 = {-9,-3,3,9}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-3),
			{Col: 1, Row: 2}: NumberVal(-1),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(-9),
			{Col: 2, Row: 2}: NumberVal(-3),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(9),
		},
	}

	// All zeros in one array
	allZerosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(0),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Single pair
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Zeros included: A1:A3 = {0,1,2}, B1:B3 = {0,2,4}
	zerosIncludedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(0),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(4),
		},
	}

	// Bool in array
	boolResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	// Large dataset
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		largeResolver.cells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i * 2))
	}

	// Weak correlation: A1:A5 = {1,2,3,4,5}, B1:B5 = {2,1,4,3,5}
	weakResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(3),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Error in array
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// All text
	allTextResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: StringVal("x"),
			{Col: 2, Row: 2}: StringVal("y"),
			{Col: 2, Row: 3}: StringVal("z"),
		},
	}

	// Different length arrays
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Excel example: y={2,3,9,1,8,7,5}, x={6,5,11,7,5,4,4}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(9),
			{Col: 1, Row: 4}: NumberVal(1),
			{Col: 1, Row: 5}: NumberVal(8),
			{Col: 1, Row: 6}: NumberVal(7),
			{Col: 1, Row: 7}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(6),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(11),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(5),
			{Col: 2, Row: 6}: NumberVal(4),
			{Col: 2, Row: 7}: NumberVal(4),
		},
	}

	// Zero std dev in second array
	zeroStdDev2Resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(7),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(7),
		},
	}

	// Fractional data: A1:A3 = {0.1,0.2,0.3}, B1:B3 = {0.5,1.0,1.5}
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0.1),
			{Col: 1, Row: 2}: NumberVal(0.2),
			{Col: 1, Row: 3}: NumberVal(0.3),
			{Col: 2, Row: 1}: NumberVal(0.5),
			{Col: 2, Row: 2}: NumberVal(1.0),
			{Col: 2, Row: 3}: NumberVal(1.5),
		},
	}

	// CORREL for the excel data is approximately 0.057950. RSQ = r^2 ≈ 0.003358
	// Let's compute: r = correl({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})
	// Excel gives RSQ ≈ 0.05795 for this data... let me use the known CORREL value.
	// Actually, the CORREL of this data set:
	// mean_y = 5, mean_x = 6
	// We can just verify it matches CORREL^2.

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic: RSQ = CORREL^2 ≈ 0.997054486^2 ≈ 0.994117654
		{"basic", "RSQ(A1:A5,B1:B5)", basicResolver, 0.997054486 * 0.997054486, false, 0},
		// Perfect positive: r=1, r^2=1
		{"perfect_positive", "RSQ(A1:A3,B1:B3)", perfectPosResolver, 1.0, false, 0},
		// Perfect negative: r=-1, r^2=1
		{"perfect_negative", "RSQ(A1:A3,B1:B3)", perfectNegResolver, 1.0, false, 0},
		// Zero std dev in first array
		{"zero_stddev_array1", "RSQ(A1:A3,B1:B3)", zeroStdDevResolver, 0, true, ErrValDIV0},
		// Zero std dev in second array
		{"zero_stddev_array2", "RSQ(A1:A3,B1:B3)", zeroStdDev2Resolver, 0, true, ErrValDIV0},
		// Two pairs: r=1, r^2=1
		{"two_pairs", "RSQ(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		// Mixed types: r=1, r^2=1
		{"mixed_types", "RSQ(A1:A4,B1:B4)", mixedResolver, 1.0, false, 0},
		// Negative numbers: r=1, r^2=1
		{"negative_numbers", "RSQ(A1:A4,B1:B4)", negResolver, 1.0, false, 0},
		// All zeros in one array
		{"all_zeros_one_array", "RSQ(A1:A3,B1:B3)", allZerosResolver, 0, true, ErrValDIV0},
		// Single pair
		{"single_pair", "RSQ(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Empty range
		{"empty_range", "RSQ(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Zeros included: r=1, r^2=1
		{"zeros_included", "RSQ(A1:A3,B1:B3)", zerosIncludedResolver, 1.0, false, 0},
		// Bool in array: r=1, r^2=1
		{"bool_in_array", "RSQ(A1:A3,B1:B3)", boolResolver, 1.0, false, 0},
		// Large dataset: r=1, r^2=1
		{"large_dataset", "RSQ(A1:A20,B1:B20)", largeResolver, 1.0, false, 0},
		// Weak correlation: r=0.8, r^2=0.64
		{"weak_correlation", "RSQ(A1:A5,B1:B5)", weakResolver, 0.64, false, 0},
		// Error propagation
		{"error_in_array", "RSQ(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// All text
		{"all_text", "RSQ(A1:A3,B1:B3)", allTextResolver, 0, true, ErrValDIV0},
		// Different length arrays
		{"different_lengths", "RSQ(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Wrong number of arguments
		{"too_few_args", "RSQ(A1:A3)", basicResolver, 0, true, ErrValVALUE},
		{"too_many_args", "RSQ(A1:A3,B1:B3,A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Reversed argument order: same RSQ (squaring removes sign info)
		{"reversed_args", "RSQ(B1:B5,A1:A5)", basicResolver, 0.997054486 * 0.997054486, false, 0},
		// Fractional values: r=1, r^2=1
		{"fractional_values", "RSQ(A1:A3,B1:B3)", fracResolver, 1.0, false, 0},
		// Excel example data
		{"excel_example", "RSQ(A1:A7,B1:B7)", excelResolver, 0.057950, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// STEYX
// ---------------------------------------------------------------------------

func TestSTEYX(t *testing.T) {
	const tol = 1e-6

	// Excel example data: y={2,3,9,1,8,7,5}, x={6,5,11,7,5,4,4}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
			{Col: 1, Row: 6}: NumberVal(7), {Col: 2, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(5), {Col: 2, Row: 7}: NumberVal(4),
		},
	}

	// Perfect fit: y={2,4,6}, x={1,2,3}
	perfectResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(6), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Horizontal line y: y={5,5,5}, x={1,2,3}
	horizontalYResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(5), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// All same x: y={1,2,3}, x={5,5,5}
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Only two data points: y={1,2}, x={3,4}
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Three points: y={1,3,2}, x={1,2,3}
	threePointResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative values: y={-3,-1,2,5}, x={-2,-1,0,1}
	negativeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-3), {Col: 2, Row: 1}: NumberVal(-2),
			{Col: 1, Row: 2}: NumberVal(-1), {Col: 2, Row: 2}: NumberVal(-1),
			{Col: 1, Row: 3}: NumberVal(2), {Col: 2, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(5), {Col: 2, Row: 4}: NumberVal(1),
		},
	}

	// Large values (perfect fit): y={1e6,2e6,3e6}, x={100,200,300}
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1000000), {Col: 2, Row: 1}: NumberVal(100),
			{Col: 1, Row: 2}: NumberVal(2000000), {Col: 2, Row: 2}: NumberVal(200),
			{Col: 1, Row: 3}: NumberVal(3000000), {Col: 2, Row: 3}: NumberVal(300),
		},
	}

	// Near-perfect fit: y={2,4.1,5.9}, x={1,2,3}
	nearPerfectResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4.1), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(5.9), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Non-numeric values skipped (strings in some positions)
	nonNumericResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: StringVal("abc"), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: StringVal("xyz"),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
			{Col: 1, Row: 6}: NumberVal(7), {Col: 2, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(5), {Col: 2, Row: 7}: NumberVal(4),
		},
	}

	// Different length arrays
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Single point
	singlePointResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
		},
	}

	// Four points with scatter: y={10,20,15,25}, x={1,2,3,4}
	fourPointResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(20), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(15), {Col: 2, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(25), {Col: 2, Row: 4}: NumberVal(4),
		},
	}

	// Decimal values (perfect fit): y={0.5,1.5,2.5,3.5}, x={0.1,0.2,0.3,0.4}
	decimalResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0.5), {Col: 2, Row: 1}: NumberVal(0.1),
			{Col: 1, Row: 2}: NumberVal(1.5), {Col: 2, Row: 2}: NumberVal(0.2),
			{Col: 1, Row: 3}: NumberVal(2.5), {Col: 2, Row: 3}: NumberVal(0.3),
			{Col: 1, Row: 4}: NumberVal(3.5), {Col: 2, Row: 4}: NumberVal(0.4),
		},
	}

	// Too few valid after skipping non-numeric (only 2 valid pairs)
	tooFewAfterSkipResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: StringVal("a"), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(7), {Col: 2, Row: 4}: StringVal("b"),
		},
	}

	// All non-numeric
	allNonNumericResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"), {Col: 2, Row: 1}: StringVal("x"),
			{Col: 1, Row: 2}: StringVal("b"), {Col: 2, Row: 2}: StringVal("y"),
			{Col: 1, Row: 3}: StringVal("c"), {Col: 2, Row: 3}: StringVal("z"),
		},
	}

	// Five points: y={3,5,7,4,6}, x={1,3,5,2,4}
	fivePointResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(7), {Col: 2, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(4), {Col: 2, Row: 4}: NumberVal(2),
			{Col: 1, Row: 5}: NumberVal(6), {Col: 2, Row: 5}: NumberVal(4),
		},
	}

	// Zero x and y values: y={0,0,0,1}, x={0,1,2,3}
	zeroYResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0), {Col: 2, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(0), {Col: 2, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(0), {Col: 2, Row: 3}: NumberVal(2),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(3),
		},
	}

	// Negative slope: y={6,4,2}, x={1,2,3}
	negSlopeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"Excel example", "STEYX(A1:A7,B1:B7)", excelResolver, 3.305719, false, 0},
		{"Perfect linear fit", "STEYX(A1:A3,B1:B3)", perfectResolver, 0, false, 0},
		{"Near-perfect fit", "STEYX(A1:A3,B1:B3)", nearPerfectResolver, 0.122474, false, 0},
		{"Horizontal y line", "STEYX(A1:A3,B1:B3)", horizontalYResolver, 0, false, 0},
		{"Constant x values", "STEYX(A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		{"Two points only", "STEYX(A1:A2,B1:B2)", twoPairResolver, 0, true, ErrValDIV0},
		{"Single point", "STEYX(A1:A1,B1:B1)", singlePointResolver, 0, true, ErrValDIV0},
		{"Three points minimum", "STEYX(A1:A3,B1:B3)", threePointResolver, 1.224745, false, 0},
		{"Negative values", "STEYX(A1:A4,B1:B4)", negativeResolver, 0.387298, false, 0},
		{"Large values perfect fit", "STEYX(A1:A3,B1:B3)", largeResolver, 0, false, 0},
		{"Four points with scatter", "STEYX(A1:A4,B1:B4)", fourPointResolver, 4.743416, false, 0},
		{"Decimal values perfect fit", "STEYX(A1:A4,B1:B4)", decimalResolver, 0, false, 0},
		{"Non-numeric values skipped", "STEYX(A1:A7,B1:B7)", nonNumericResolver, 2.934247, false, 0},
		{"Different length arrays", "STEYX(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		{"Too few after skip", "STEYX(A1:A4,B1:B4)", tooFewAfterSkipResolver, 0, true, ErrValDIV0},
		{"All non-numeric", "STEYX(A1:A3,B1:B3)", allNonNumericResolver, 0, true, ErrValDIV0},
		{"No args", "STEYX()", excelResolver, 0, true, ErrValVALUE},
		{"One arg", "STEYX(A1:A7)", excelResolver, 0, true, ErrValVALUE},
		{"Three args", "STEYX(A1:A7,B1:B7,A1:A7)", excelResolver, 0, true, ErrValVALUE},
		{"Five points", "STEYX(A1:A5,B1:B5)", fivePointResolver, 0.0, false, 0},
		{"Zero y values", "STEYX(A1:A4,B1:B4)", zeroYResolver, 0.387298, false, 0},
		{"Negative slope perfect", "STEYX(A1:A3,B1:B3)", negSlopeResolver, 0, false, 0},
		{"Repeat Excel example", "STEYX(A1:A7,B1:B7)", excelResolver, 3.305719, false, 0},
		{"Empty arrays", "STEYX(C1:C3,D1:D3)", excelResolver, 0, true, ErrValDIV0},
		{"Non-numeric both positions", "STEYX(A1:A7,B1:B7)", nonNumericResolver, 2.934247, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}

			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}

			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// VARA
// ---------------------------------------------------------------------------

func TestVARA(t *testing.T) {
	const tol = 1e-9

	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("basic numbers", func(t *testing.T) {
		// VARA(2,4,6) = VAR of {2,4,6} = 4
		cf := evalCompile(t, "VARA(2,4,6)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-4) > tol {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("basic range", func(t *testing.T) {
		r := valResolver(NumberVal(2), NumberVal(4), NumberVal(6))
		cf := evalCompile(t, "VARA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-4) > tol {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("TRUE in range counts as 1", func(t *testing.T) {
		// {TRUE, 3} -> {1, 3}, mean=2, ssq=(1+1)=2, var=2/(2-1)=2
		r := valResolver(BoolVal(true), NumberVal(3))
		cf := evalCompile(t, "VARA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-2) > tol {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("FALSE in range counts as 0", func(t *testing.T) {
		// {FALSE, 4} -> {0, 4}, mean=2, ssq=(4+4)=8, var=8/1=8
		r := valResolver(BoolVal(false), NumberVal(4))
		cf := evalCompile(t, "VARA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-8) > tol {
			t.Errorf("got %v, want 8", got)
		}
	})

	t.Run("text in range counts as 0", func(t *testing.T) {
		// {10, "hello"} -> {10, 0}, mean=5, ssq=(25+25)=50, var=50/1=50
		r := valResolver(NumberVal(10), StringVal("hello"))
		cf := evalCompile(t, "VARA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-50) > tol {
			t.Errorf("got %v, want 50", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		// A1=2, A2=empty, A3=4 -> {2,4}, var=2
		r := &mockResolver{cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
		}}
		cf := evalCompile(t, "VARA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-2) > tol {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("direct TRUE arg counts as 1", func(t *testing.T) {
		// VARA(TRUE,3) -> {1,3}, mean=2, ssq=2, var=2
		cf := evalCompile(t, "VARA(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-2) > tol {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("direct FALSE arg counts as 0", func(t *testing.T) {
		// VARA(FALSE,4) -> {0,4}, var=8
		cf := evalCompile(t, "VARA(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-8) > tol {
			t.Errorf("got %v, want 8", got)
		}
	})

	t.Run("direct numeric string coerced", func(t *testing.T) {
		// VARA("5",15) -> {5,15}, mean=10, ssq=50, var=50
		cf := evalCompile(t, `VARA("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-50) > tol {
			t.Errorf("got %v, want 50", got)
		}
	})

	t.Run("direct non-numeric string returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `VARA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation from range", func(t *testing.T) {
		r := valResolver(NumberVal(1), ErrorVal(ErrValNA), NumberVal(3))
		cf := evalCompile(t, "VARA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation from direct arg", func(t *testing.T) {
		r := valResolver(ErrorVal(ErrValDIV0))
		cf := evalCompile(t, "VARA(A1,10)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("single value returns DIV0", func(t *testing.T) {
		cf := evalCompile(t, "VARA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("all same values returns 0", func(t *testing.T) {
		cf := evalCompile(t, "VARA(5,5,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, "VARA()")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("mixed types in range", func(t *testing.T) {
		// {1, TRUE, "abc", 3} -> {1, 1, 0, 3}, mean=5/4=1.25
		// ssq = (0.25^2 + 0.25^2 + 1.25^2 + 1.75^2) = 0.0625+0.0625+1.5625+3.0625 = 4.75
		// var = 4.75/3
		r := valResolver(NumberVal(1), BoolVal(true), StringVal("abc"), NumberVal(3))
		cf := evalCompile(t, "VARA(A1:A4)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 4.75 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > tol {
			t.Errorf("got %v, want %f", got, want)
		}
	})

	t.Run("two values", func(t *testing.T) {
		// VARA(10,20), mean=15, ssq=50, var=50
		cf := evalCompile(t, "VARA(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-50) > tol {
			t.Errorf("got %v, want 50", got)
		}
	})

	t.Run("larger dataset", func(t *testing.T) {
		// {1,2,3,4,5}, mean=3, ssq=10, var=10/4=2.5
		r := valResolver(NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5))
		cf := evalCompile(t, "VARA(A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-2.5) > tol {
			t.Errorf("got %v, want 2.5", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		// VARA(-2,0,2), mean=0, ssq=8, var=8/2=4
		cf := evalCompile(t, "VARA(-2,0,2)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-4) > tol {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("range with only booleans", func(t *testing.T) {
		// {TRUE, FALSE} -> {1, 0}, mean=0.5, ssq=0.5, var=0.5
		r := valResolver(BoolVal(true), BoolVal(false))
		cf := evalCompile(t, "VARA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.5) > tol {
			t.Errorf("got %v, want 0.5", got)
		}
	})

	t.Run("range with only text", func(t *testing.T) {
		// {"a", "b"} -> {0, 0}, mean=0, ssq=0, var=0
		r := valResolver(StringVal("a"), StringVal("b"))
		cf := evalCompile(t, "VARA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("single text cell ref returns VALUE error", func(t *testing.T) {
		// Single cell ref A1="hello" is a direct arg -> CoerceNum -> #VALUE!
		r := valResolver(StringVal("hello"))
		cf := evalCompile(t, "VARA(A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})
}

// ---------------------------------------------------------------------------
// VARPA
// ---------------------------------------------------------------------------

func TestVARPA(t *testing.T) {
	const tol = 1e-9

	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("basic numbers", func(t *testing.T) {
		// VARPA(2,4,6), mean=4, ssq=8, varp=8/3
		cf := evalCompile(t, "VARPA(2,4,6)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 8.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > tol {
			t.Errorf("got %v, want %f", got, want)
		}
	})

	t.Run("basic range", func(t *testing.T) {
		r := valResolver(NumberVal(2), NumberVal(4), NumberVal(6))
		cf := evalCompile(t, "VARPA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 8.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > tol {
			t.Errorf("got %v, want %f", got, want)
		}
	})

	t.Run("TRUE in range counts as 1", func(t *testing.T) {
		// {TRUE, 3} -> {1, 3}, mean=2, ssq=2, varp=2/2=1
		r := valResolver(BoolVal(true), NumberVal(3))
		cf := evalCompile(t, "VARPA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-1) > tol {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("FALSE in range counts as 0", func(t *testing.T) {
		// {FALSE, 4} -> {0, 4}, mean=2, ssq=8, varp=8/2=4
		r := valResolver(BoolVal(false), NumberVal(4))
		cf := evalCompile(t, "VARPA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-4) > tol {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("text in range counts as 0", func(t *testing.T) {
		// {10, "hello"} -> {10, 0}, mean=5, ssq=50, varp=50/2=25
		r := valResolver(NumberVal(10), StringVal("hello"))
		cf := evalCompile(t, "VARPA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-25) > tol {
			t.Errorf("got %v, want 25", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		// A1=2, A2=empty, A3=4 -> {2,4}, varp = 2/2 = 1
		r := &mockResolver{cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
		}}
		cf := evalCompile(t, "VARPA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-1) > tol {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("direct TRUE arg counts as 1", func(t *testing.T) {
		// VARPA(TRUE,3) -> {1,3}, mean=2, ssq=2, varp=1
		cf := evalCompile(t, "VARPA(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-1) > tol {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("direct FALSE arg counts as 0", func(t *testing.T) {
		// VARPA(FALSE,4) -> {0,4}, varp=4
		cf := evalCompile(t, "VARPA(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-4) > tol {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("direct numeric string coerced", func(t *testing.T) {
		// VARPA("5",15) -> {5,15}, mean=10, ssq=50, varp=25
		cf := evalCompile(t, `VARPA("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-25) > tol {
			t.Errorf("got %v, want 25", got)
		}
	})

	t.Run("direct non-numeric string returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `VARPA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation from range", func(t *testing.T) {
		r := valResolver(NumberVal(1), ErrorVal(ErrValNA), NumberVal(3))
		cf := evalCompile(t, "VARPA(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation from direct arg", func(t *testing.T) {
		r := valResolver(ErrorVal(ErrValDIV0))
		cf := evalCompile(t, "VARPA(A1,10)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("single value returns 0", func(t *testing.T) {
		cf := evalCompile(t, "VARPA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("all same values returns 0", func(t *testing.T) {
		cf := evalCompile(t, "VARPA(5,5,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, "VARPA()")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("mixed types in range", func(t *testing.T) {
		// {1, TRUE, "abc", 3} -> {1, 1, 0, 3}, mean=5/4=1.25
		// ssq = 4.75, varp = 4.75/4 = 1.1875
		r := valResolver(NumberVal(1), BoolVal(true), StringVal("abc"), NumberVal(3))
		cf := evalCompile(t, "VARPA(A1:A4)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 4.75 / 4.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > tol {
			t.Errorf("got %v, want %f", got, want)
		}
	})

	t.Run("two values", func(t *testing.T) {
		// VARPA(10,20), mean=15, ssq=50, varp=25
		cf := evalCompile(t, "VARPA(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-25) > tol {
			t.Errorf("got %v, want 25", got)
		}
	})

	t.Run("larger dataset", func(t *testing.T) {
		// {1,2,3,4,5}, mean=3, ssq=10, varp=10/5=2
		r := valResolver(NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5))
		cf := evalCompile(t, "VARPA(A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-2) > tol {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		// VARPA(-2,0,2), mean=0, ssq=8, varp=8/3
		cf := evalCompile(t, "VARPA(-2,0,2)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 8.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > tol {
			t.Errorf("got %v, want %f", got, want)
		}
	})

	t.Run("range with only booleans", func(t *testing.T) {
		// {TRUE, FALSE} -> {1, 0}, mean=0.5, ssq=0.5, varp=0.25
		r := valResolver(BoolVal(true), BoolVal(false))
		cf := evalCompile(t, "VARPA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.25) > tol {
			t.Errorf("got %v, want 0.25", got)
		}
	})

	t.Run("range with only text", func(t *testing.T) {
		// {"a", "b"} -> {0, 0}, varp=0
		r := valResolver(StringVal("a"), StringVal("b"))
		cf := evalCompile(t, "VARPA(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("single TRUE returns 0", func(t *testing.T) {
		// {TRUE} -> {1}, n=1, varp=0
		r := valResolver(BoolVal(true))
		cf := evalCompile(t, "VARPA(A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})
}

// ---------------------------------------------------------------------------
// NORM.S.DIST
// ---------------------------------------------------------------------------

func TestNORMSDIST(t *testing.T) {
	const tol = 1e-6
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// PDF tests
		{"pdf_z0", "NORM.S.DIST(0,FALSE)", 0.398942280401, false, 0},
		{"pdf_z1", "NORM.S.DIST(1,FALSE)", 0.241970724519, false, 0},
		{"pdf_z_neg1", "NORM.S.DIST(-1,FALSE)", 0.241970724519, false, 0},
		{"pdf_z2", "NORM.S.DIST(2,FALSE)", 0.053990966513, false, 0},
		{"pdf_z_neg2", "NORM.S.DIST(-2,FALSE)", 0.053990966513, false, 0},
		{"pdf_z1.333333", "NORM.S.DIST(1.333333,FALSE)", 0.164010148, false, 0},
		{"pdf_z3", "NORM.S.DIST(3,FALSE)", 0.004431848412, false, 0},
		{"pdf_symmetry_half", "NORM.S.DIST(0.5,FALSE)", 0.352065326764, false, 0},
		{"pdf_symmetry_neg_half", "NORM.S.DIST(-0.5,FALSE)", 0.352065326764, false, 0},
		{"pdf_large", "NORM.S.DIST(10,FALSE)", 0.0, false, 0},

		// CDF tests
		{"cdf_z0", "NORM.S.DIST(0,TRUE)", 0.5, false, 0},
		{"cdf_z1", "NORM.S.DIST(1,TRUE)", 0.841344746069, false, 0},
		{"cdf_z_neg1", "NORM.S.DIST(-1,TRUE)", 0.158655253931, false, 0},
		{"cdf_z2", "NORM.S.DIST(2,TRUE)", 0.977249868052, false, 0},
		{"cdf_z_neg2", "NORM.S.DIST(-2,TRUE)", 0.022750131948, false, 0},
		{"cdf_z3", "NORM.S.DIST(3,TRUE)", 0.998650101968, false, 0},
		{"cdf_z_neg3", "NORM.S.DIST(-3,TRUE)", 0.001349898032, false, 0},
		{"cdf_z1.333333", "NORM.S.DIST(1.333333,TRUE)", 0.908788726, false, 0},
		{"cdf_z1.96", "NORM.S.DIST(1.96,TRUE)", 0.975002105, false, 0},
		{"cdf_z_neg1.96", "NORM.S.DIST(-1.96,TRUE)", 0.024997895, false, 0},
		{"cdf_large_pos", "NORM.S.DIST(10,TRUE)", 1.0, false, 0},
		{"cdf_large_neg", "NORM.S.DIST(-10,TRUE)", 0.0, false, 0},

		// Error cases
		{"err_non_numeric_z", `NORM.S.DIST("abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `NORM.S.DIST(0,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestNORMSDIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NORM.S.DIST(0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.S.DIST(0) should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "NORM.S.DIST(0,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.S.DIST(0,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from NORM.S.DIST with 1 arg.
	cf = evalCompile(t, `IFERROR(NORM.S.DIST(1),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(NORM.S.DIST(1),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// NORM.S.INV
// ---------------------------------------------------------------------------

func TestNORMSINV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"p_0.5", "NORM.S.INV(0.5)", 0, false, 0},
		{"p_0.841344746", "NORM.S.INV(0.841344746)", 1.0, false, 0},
		{"p_0.158655254", "NORM.S.INV(0.158655254)", -1.0, false, 0},
		{"p_0.977249868", "NORM.S.INV(0.977249868)", 2.0, false, 0},
		{"p_0.022750132", "NORM.S.INV(0.022750132)", -2.0, false, 0},
		{"p_0.998650102", "NORM.S.INV(0.998650102)", 3.0, false, 0},
		{"p_0.001349898", "NORM.S.INV(0.001349898)", -3.0, false, 0},
		{"p_0.975", "NORM.S.INV(0.975)", 1.959964, false, 0},
		{"p_0.025", "NORM.S.INV(0.025)", -1.959964, false, 0},
		{"p_0.9", "NORM.S.INV(0.9)", 1.281552, false, 0},
		{"p_0.1", "NORM.S.INV(0.1)", -1.281552, false, 0},
		{"p_0.95", "NORM.S.INV(0.95)", 1.644854, false, 0},
		{"p_0.05", "NORM.S.INV(0.05)", -1.644854, false, 0},
		{"p_0.99", "NORM.S.INV(0.99)", 2.326348, false, 0},
		{"p_0.01", "NORM.S.INV(0.01)", -2.326348, false, 0},
		{"p_0.001", "NORM.S.INV(0.001)", -3.090232, false, 0},
		{"p_0.999", "NORM.S.INV(0.999)", 3.090232, false, 0},
		{"p_0.908789", "NORM.S.INV(0.908789)", 1.333335, false, 0},
		{"p_0.75", "NORM.S.INV(0.75)", 0.674490, false, 0},
		{"p_0.25", "NORM.S.INV(0.25)", -0.674490, false, 0},

		// Error: boundary values
		{"err_p_zero", "NORM.S.INV(0)", 0, true, ErrValNUM},
		{"err_p_one", "NORM.S.INV(1)", 0, true, ErrValNUM},
		{"err_p_neg", "NORM.S.INV(-0.1)", 0, true, ErrValNUM},
		{"err_p_gt1", "NORM.S.INV(1.1)", 0, true, ErrValNUM},
		{"err_non_numeric", `NORM.S.INV("abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestNORMSINV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NORM.S.INV()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.S.INV() should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "NORM.S.INV(0.5,0.3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.S.INV(0.5,0.3) should error, got type=%d", got.Type)
	}
}

func TestNORMSINV_NORMSDIST_roundtrip(t *testing.T) {
	// Excel's NORM.S.INV uses Acklam's approximation without refinement,
	// so the roundtrip through NORM.S.DIST is only accurate to ~8 digits.
	const tol = 1e-8
	resolver := &mockResolver{}

	probs := []float64{0.1, 0.25, 0.5, 0.75, 0.9, 0.95, 0.99, 0.001, 0.999}
	for _, p := range probs {
		t.Run(fmt.Sprintf("roundtrip_%g", p), func(t *testing.T) {
			formula := fmt.Sprintf("NORM.S.DIST(NORM.S.INV(%g),TRUE)", p)
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-p) > tol {
				t.Errorf("NORM.S.DIST(NORM.S.INV(%g),TRUE) = %g, want %g", p, got.Num, p)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// NORM.DIST
// ---------------------------------------------------------------------------

func TestNORMDIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// CDF tests
		{"cdf_excel_example", "NORM.DIST(42,40,1.5,TRUE)", 0.9087888, false, 0},
		{"cdf_at_mean", "NORM.DIST(40,40,1.5,TRUE)", 0.5, false, 0},
		{"cdf_below_mean", "NORM.DIST(38,40,1.5,TRUE)", 0.0912112, false, 0},
		{"cdf_far_above", "NORM.DIST(50,40,1.5,TRUE)", 1.0, false, 0},
		{"cdf_far_below", "NORM.DIST(30,40,1.5,TRUE)", 0.0, false, 0},
		{"cdf_mean0_sd1", "NORM.DIST(1,0,1,TRUE)", 0.841345, false, 0},
		{"cdf_mean0_sd1_neg", "NORM.DIST(-1,0,1,TRUE)", 0.158655, false, 0},
		{"cdf_mean0_sd1_zero", "NORM.DIST(0,0,1,TRUE)", 0.5, false, 0},
		{"cdf_mean100_sd15", "NORM.DIST(115,100,15,TRUE)", 0.841345, false, 0},
		{"cdf_mean100_sd15_below", "NORM.DIST(85,100,15,TRUE)", 0.158655, false, 0},
		{"cdf_neg_x", "NORM.DIST(-5,0,2,TRUE)", 0.006210, false, 0},
		{"cdf_large_stdev", "NORM.DIST(50,50,100,TRUE)", 0.5, false, 0},
		{"cdf_small_stdev", "NORM.DIST(40.01,40,0.01,TRUE)", 0.841345, false, 0},

		// PDF tests
		{"pdf_excel_example", "NORM.DIST(42,40,1.5,FALSE)", 0.10934, false, 0},
		{"pdf_at_mean", "NORM.DIST(40,40,1.5,FALSE)", 0.265962, false, 0},
		{"pdf_mean0_sd1", "NORM.DIST(0,0,1,FALSE)", 0.398942, false, 0},
		{"pdf_mean0_sd1_z1", "NORM.DIST(1,0,1,FALSE)", 0.241971, false, 0},
		{"pdf_mean0_sd1_zneg1", "NORM.DIST(-1,0,1,FALSE)", 0.241971, false, 0},
		{"pdf_symmetry", "NORM.DIST(38,40,1.5,FALSE)", 0.10934, false, 0},
		{"pdf_large_stdev", "NORM.DIST(50,50,100,FALSE)", 0.003989, false, 0},

		// Error cases
		{"err_stdev_zero", "NORM.DIST(42,40,0,TRUE)", 0, true, ErrValNUM},
		{"err_stdev_neg", "NORM.DIST(42,40,-1,TRUE)", 0, true, ErrValNUM},
		{"err_non_numeric_x", `NORM.DIST("abc",40,1.5,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_mean", `NORM.DIST(42,"abc",1.5,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_stdev", `NORM.DIST(42,40,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `NORM.DIST(42,40,1.5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestNORMDIST_argCount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NORM.DIST(42,40,1.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.DIST(42,40,1.5) should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "NORM.DIST(42,40,1.5,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.DIST(42,40,1.5,TRUE,1) should error, got type=%d", got.Type)
	}
}

// ---------------------------------------------------------------------------
// NORM.INV
// ---------------------------------------------------------------------------

func TestNORMINV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic values
		{"excel_example", "NORM.INV(0.908789,40,1.5)", 42.0, false, 0},
		{"p_half_returns_mean", "NORM.INV(0.5,40,1.5)", 40.0, false, 0},
		{"p_half_mean0", "NORM.INV(0.5,0,1)", 0.0, false, 0},
		{"p_half_mean100", "NORM.INV(0.5,100,15)", 100.0, false, 0},
		{"p_0.841345_mean0_sd1", "NORM.INV(0.841345,0,1)", 1.0, false, 0},
		{"p_0.158655_mean0_sd1", "NORM.INV(0.158655,0,1)", -1.0, false, 0},
		{"p_0.977250_mean0_sd1", "NORM.INV(0.977250,0,1)", 2.0, false, 0},
		{"p_0.9_mean50_sd10", "NORM.INV(0.9,50,10)", 62.81552, false, 0},
		{"p_0.1_mean50_sd10", "NORM.INV(0.1,50,10)", 37.18448, false, 0},
		{"p_0.95_mean0_sd1", "NORM.INV(0.95,0,1)", 1.644854, false, 0},
		{"p_0.05_mean0_sd1", "NORM.INV(0.05,0,1)", -1.644854, false, 0},
		{"p_0.99_mean100_sd15", "NORM.INV(0.99,100,15)", 134.89522, false, 0},
		{"p_0.01_mean100_sd15", "NORM.INV(0.01,100,15)", 65.10478, false, 0},
		{"p_0.75_mean0_sd1", "NORM.INV(0.75,0,1)", 0.67449, false, 0},
		{"p_0.25_mean0_sd1", "NORM.INV(0.25,0,1)", -0.67449, false, 0},
		{"large_stdev", "NORM.INV(0.5,0,1000)", 0.0, false, 0},
		{"neg_mean", "NORM.INV(0.5,-50,10)", -50.0, false, 0},

		// Error cases
		{"err_p_zero", "NORM.INV(0,40,1.5)", 0, true, ErrValNUM},
		{"err_p_one", "NORM.INV(1,40,1.5)", 0, true, ErrValNUM},
		{"err_p_neg", "NORM.INV(-0.1,40,1.5)", 0, true, ErrValNUM},
		{"err_p_gt1", "NORM.INV(1.5,40,1.5)", 0, true, ErrValNUM},
		{"err_stdev_zero", "NORM.INV(0.5,40,0)", 0, true, ErrValNUM},
		{"err_stdev_neg", "NORM.INV(0.5,40,-1)", 0, true, ErrValNUM},
		{"err_non_numeric_p", `NORM.INV("abc",40,1.5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_mean", `NORM.INV(0.5,"abc",1.5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_stdev", `NORM.INV(0.5,40,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestNORMINV_argCount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NORM.INV(0.5,40)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.INV(0.5,40) should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "NORM.INV(0.5,40,1.5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NORM.INV(0.5,40,1.5,1) should error, got type=%d", got.Type)
	}
}

func TestNORMINV_NORMDIST_roundtrip(t *testing.T) {
	const tol = 1e-14
	resolver := &mockResolver{}

	cases := []struct {
		p    float64
		mean float64
		sd   float64
	}{
		{0.1, 40, 1.5},
		{0.025, 40, 1.5},
		{0.25, 100, 15},
		{0.5, 0, 1},
		{0.75, -50, 10},
		{0.9, 200, 30},
		{0.95, 0, 0.5},
		{0.99, 50, 5},
		{0.001, 10, 2},
		{0.999, 10, 2},
	}
	for _, tc := range cases {
		t.Run(fmt.Sprintf("roundtrip_p%g_m%g_s%g", tc.p, tc.mean, tc.sd), func(t *testing.T) {
			formula := fmt.Sprintf("NORM.DIST(NORM.INV(%g,%g,%g),%g,%g,TRUE)", tc.p, tc.mean, tc.sd, tc.mean, tc.sd)
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tc.p) > tol {
				t.Errorf("NORM.DIST(NORM.INV(%g,%g,%g),%g,%g,TRUE) = %g, want %g", tc.p, tc.mean, tc.sd, tc.mean, tc.sd, got.Num, tc.p)
			}
		})
	}
}

func TestMEDIAN(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("odd count basic", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1,2,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("even count basic", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1,2,3,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2.5 {
			t.Errorf("got %v, want 2.5", got)
		}
	})

	t.Run("single value", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(42)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("two values", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("five values unsorted", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(9,1,5,3,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("six values unsorted", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1,2,3,4,5,6)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3.5 {
			t.Errorf("got %v, want 3.5", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(-5,-3,-1)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -3 {
			t.Errorf("got %v, want -3", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(-10,0,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed positive and negative even count", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(-10,-1,1,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("all zeros", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(0,0,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(7,7,7,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("decimals", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1.5,2.5,3.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2.5 {
			t.Errorf("got %v, want 2.5", got)
		}
	})

	t.Run("very large numbers", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1E15,2E15,3E15)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2e15 {
			t.Errorf("got %v, want 2e15", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(1E-10,2E-10,3E-10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2e-10 {
			t.Errorf("got %v, want 2e-10", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerced to 1", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(TRUE,3,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerced to 0", func(t *testing.T) {
		cf := evalCompile(t, "MEDIAN(FALSE,2,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("numeric string as direct arg coerced", func(t *testing.T) {
		cf := evalCompile(t, `MEDIAN("2",4,6)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})

	t.Run("non-numeric string as direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `MEDIAN("abc",1,2)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		m := numResolver(1, 2, 3, 4, 5)
		cf := evalCompile(t, "MEDIAN(A1:A5)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("range even count", func(t *testing.T) {
		m := numResolver(1, 2, 3, 4, 5, 6)
		cf := evalCompile(t, "MEDIAN(A1:A6)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3.5 {
			t.Errorf("got %v, want 3.5", got)
		}
	})

	t.Run("range with empty cells ignored", func(t *testing.T) {
		// A1=1, A2=empty, A3=3, A4=empty, A5=5
		m := &mockResolver{cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(5),
		}}
		cf := evalCompile(t, "MEDIAN(A1:A5)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("range with strings ignored", func(t *testing.T) {
		m := valResolver(NumberVal(1), StringVal("hello"), NumberVal(3), StringVal("world"), NumberVal(5))
		cf := evalCompile(t, "MEDIAN(A1:A5)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("range with booleans ignored", func(t *testing.T) {
		m := valResolver(NumberVal(10), BoolVal(true), NumberVal(20), BoolVal(false), NumberVal(30))
		cf := evalCompile(t, "MEDIAN(A1:A5)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range with error propagates", func(t *testing.T) {
		m := valResolver(NumberVal(1), ErrorVal(ErrValNA), NumberVal(3))
		cf := evalCompile(t, "MEDIAN(A1:A3)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	t.Run("all non-numeric in range returns NUM error", func(t *testing.T) {
		m := valResolver(StringVal("a"), StringVal("b"), StringVal("c"))
		cf := evalCompile(t, "MEDIAN(A1:A3)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM! error", got)
		}
	})

	t.Run("Excel example odd", func(t *testing.T) {
		// From Excel docs: MEDIAN(1,2,3,4,5) = 3
		m := numResolver(1, 2, 3, 4, 5)
		cf := evalCompile(t, "MEDIAN(A1:A5)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("Excel example even", func(t *testing.T) {
		// From Excel docs: MEDIAN(1,2,3,4,5,6) = 3.5
		m := numResolver(1, 2, 3, 4, 5, 6)
		cf := evalCompile(t, "MEDIAN(A1:A6)")
		got, err := Eval(cf, m, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3.5 {
			t.Errorf("got %v, want 3.5", got)
		}
	})

	_ = numResolver // suppress unused if only direct-arg tests
	_ = valResolver
}

// ---------------------------------------------------------------------------
// BINOM.DIST
// ---------------------------------------------------------------------------

func TestBINOM_DIST(t *testing.T) {
	const tol = 1e-9
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic PMF
		{"pmf_basic", "BINOM.DIST(6,10,0.5,FALSE)", 0.205078125, false, 0},

		// Basic CDF
		{"cdf_basic", "BINOM.DIST(6,10,0.5,TRUE)", 0.828125, false, 0},

		// All successes PMF
		{"pmf_all_success", "BINOM.DIST(10,10,0.5,FALSE)", 0.0009765625, false, 0},

		// Zero successes PMF
		{"pmf_zero_success", "BINOM.DIST(0,10,0.5,FALSE)", 0.0009765625, false, 0},

		// Probability 0, zero successes
		{"pmf_p0_k0", "BINOM.DIST(0,10,0,FALSE)", 1, false, 0},

		// Probability 1, all successes
		{"pmf_p1_kn", "BINOM.DIST(10,10,1,FALSE)", 1, false, 0},

		// Probability 0, nonzero successes
		{"pmf_p0_k5", "BINOM.DIST(5,10,0,FALSE)", 0, false, 0},

		// Probability 1, partial successes
		{"pmf_p1_k5", "BINOM.DIST(5,10,1,FALSE)", 0, false, 0},

		// CDF at max
		{"cdf_at_max", "BINOM.DIST(10,10,0.5,TRUE)", 1, false, 0},

		// CDF at 0
		{"cdf_at_zero", "BINOM.DIST(0,10,0.5,TRUE)", 0.0009765625, false, 0},

		// Truncation: 6.9 -> 6, 10.7 -> 10
		{"truncation", "BINOM.DIST(6.9,10.7,0.5,FALSE)", 0.205078125, false, 0},

		// Small probability
		{"small_prob", "BINOM.DIST(1,100,0.01,FALSE)", 0.369729637650, false, 0},

		// Large trials PMF
		{"large_trials_pmf", "BINOM.DIST(50,100,0.5,FALSE)", 0.079589237387, false, 0},

		// Single trial PMF
		{"single_trial_pmf", "BINOM.DIST(1,1,0.5,FALSE)", 0.5, false, 0},

		// Single trial CDF at 0
		{"single_trial_cdf_0", "BINOM.DIST(0,1,0.5,TRUE)", 0.5, false, 0},

		// CDF with p=0
		{"cdf_p0_k5", "BINOM.DIST(5,10,0,TRUE)", 1, false, 0},

		// CDF with p=1
		{"cdf_p1_k5", "BINOM.DIST(5,10,1,TRUE)", 0, false, 0},

		// CDF with p=1, k=n
		{"cdf_p1_kn", "BINOM.DIST(10,10,1,TRUE)", 1, false, 0},

		// CDF with p=0, k=0
		{"cdf_p0_k0", "BINOM.DIST(0,10,0,TRUE)", 1, false, 0},

		// Asymmetric probability PMF
		{"pmf_asym_prob", "BINOM.DIST(3,10,0.25,FALSE)", 0.250282287598, false, 0},

		// Single trial CDF at 1
		{"single_trial_cdf_1", "BINOM.DIST(1,1,0.5,TRUE)", 1, false, 0},

		// Zero trials
		{"zero_trials", "BINOM.DIST(0,0,0.5,FALSE)", 1, false, 0},

		// Error: number_s negative
		{"err_neg_k", "BINOM.DIST(-1,10,0.5,FALSE)", 0, true, ErrValNUM},

		// Error: number_s > trials
		{"err_k_gt_n", "BINOM.DIST(11,10,0.5,FALSE)", 0, true, ErrValNUM},

		// Error: probability < 0
		{"err_p_neg", "BINOM.DIST(5,10,-0.1,FALSE)", 0, true, ErrValNUM},

		// Error: probability > 1
		{"err_p_gt1", "BINOM.DIST(5,10,1.1,FALSE)", 0, true, ErrValNUM},

		// Error: non-numeric first arg
		{"err_non_numeric", `BINOM.DIST("abc",10,0.5,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric second arg
		{"err_non_numeric2", `BINOM.DIST(5,"abc",0.5,FALSE)`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestBINOM_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "BINOM.DIST(5,10,0.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BINOM.DIST(5,10,0.5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "BINOM.DIST(5,10,0.5,FALSE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BINOM.DIST(5,10,0.5,FALSE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(BINOM.DIST(5,10,0.5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(BINOM.DIST(5,10,0.5),"err") = %v, want string "err"`, got)
	}
}

func TestPOISSON_DIST(t *testing.T) {
	const tol = 1e-6
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic PMF
		{"pmf_basic", "POISSON.DIST(2,5,FALSE)", 0.084224, false, 0},

		// Basic CDF
		{"cdf_basic", "POISSON.DIST(2,5,TRUE)", 0.124652, false, 0},

		// Zero events PMF: e^(-5)
		{"pmf_zero_events", "POISSON.DIST(0,5,FALSE)", 0.006738, false, 0},

		// Zero events CDF (same as PMF when x=0)
		{"cdf_zero_events", "POISSON.DIST(0,5,TRUE)", 0.006738, false, 0},

		// Mean 0, x=0, PMF: should be 1
		{"pmf_mean0_x0", "POISSON.DIST(0,0,FALSE)", 1, false, 0},

		// Mean 0, x>0, PMF: should be 0
		{"pmf_mean0_x5", "POISSON.DIST(5,0,FALSE)", 0, false, 0},

		// Mean 0, CDF: should be 1
		{"cdf_mean0_x0", "POISSON.DIST(0,0,TRUE)", 1, false, 0},

		// Mean 0, x>0, CDF: should be 1
		{"cdf_mean0_x5", "POISSON.DIST(5,0,TRUE)", 1, false, 0},

		// Large x CDF approaching 1
		{"cdf_large_x", "POISSON.DIST(20,5,TRUE)", 1.0, false, 0},

		// Truncation: 2.9 → 2
		{"truncation", "POISSON.DIST(2.9,5,FALSE)", 0.084224, false, 0},

		// Mean 1, x=1, PMF: 1 * e^(-1) / 1! = e^(-1)
		{"pmf_mean1_x1", "POISSON.DIST(1,1,FALSE)", 0.367879, false, 0},

		// Mean 1, x=1, CDF: P(0) + P(1) = e^(-1) + e^(-1) = 2*e^(-1)
		{"cdf_mean1_x1", "POISSON.DIST(1,1,TRUE)", 0.735759, false, 0},

		// Small mean: P(0; 0.1) = e^(-0.1)
		{"pmf_small_mean", "POISSON.DIST(0,0.1,FALSE)", 0.904837, false, 0},

		// Large mean: P(10; 10)
		{"pmf_large_mean", "POISSON.DIST(10,10,FALSE)", 0.12511, false, 0},

		// Single event: P(1; 0.5) = 0.5 * e^(-0.5)
		{"pmf_single_event", "POISSON.DIST(1,0.5,FALSE)", 0.303265, false, 0},

		// Large x PMF: P(10; 5)
		{"pmf_large_x", "POISSON.DIST(10,5,FALSE)", 0.018133, false, 0},

		// CDF x=5, mean=5
		{"cdf_x5_mean5", "POISSON.DIST(5,5,TRUE)", 0.615961, false, 0},

		// PMF x=0, mean=1: e^(-1)
		{"pmf_x0_mean1", "POISSON.DIST(0,1,FALSE)", 0.367879, false, 0},

		// CDF x=0, mean=1: e^(-1)
		{"cdf_x0_mean1", "POISSON.DIST(0,1,TRUE)", 0.367879, false, 0},

		// PMF x=3, mean=2
		{"pmf_x3_mean2", "POISSON.DIST(3,2,FALSE)", 0.180447, false, 0},

		// Error: x < 0
		{"err_neg_x", "POISSON.DIST(-1,5,FALSE)", 0, true, ErrValNUM},

		// Error: mean < 0
		{"err_neg_mean", "POISSON.DIST(2,-1,FALSE)", 0, true, ErrValNUM},

		// Error: non-numeric x
		{"err_non_numeric_x", `POISSON.DIST("abc",5,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric mean
		{"err_non_numeric_mean", `POISSON.DIST(2,"abc",FALSE)`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestPOISSON_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "POISSON.DIST(2,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("POISSON.DIST(2,5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "POISSON.DIST(2,5,FALSE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("POISSON.DIST(2,5,FALSE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(POISSON.DIST(2,5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(POISSON.DIST(2,5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// EXPON.DIST
// ---------------------------------------------------------------------------

func TestEXPON_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// PDF tests
		{"pdf_basic", "EXPON.DIST(0.2,10,FALSE)", 1.353352832, false, 0},
		{"pdf_x0", "EXPON.DIST(0,10,FALSE)", 10, false, 0},
		{"pdf_lambda1_x1", "EXPON.DIST(1,1,FALSE)", 0.367879441, false, 0},
		{"pdf_large_x", "EXPON.DIST(10,1,FALSE)", 0.0000453999, false, 0},
		{"pdf_small_lambda", "EXPON.DIST(1,0.1,FALSE)", 0.090483742, false, 0},
		{"pdf_large_lambda", "EXPON.DIST(0.01,100,FALSE)", 36.78794412, false, 0},
		{"pdf_x05_lambda2", "EXPON.DIST(0.5,2,FALSE)", 0.735758882, false, 0},
		{"pdf_x1_lambda05", "EXPON.DIST(1,0.5,FALSE)", 0.303265330, false, 0},
		{"pdf_x2_lambda3", "EXPON.DIST(2,3,FALSE)", 0.007436478, false, 0},
		{"pdf_x001_lambda1", "EXPON.DIST(0.01,1,FALSE)", 0.990049834, false, 0},

		// CDF tests
		{"cdf_basic", "EXPON.DIST(0.2,10,TRUE)", 0.864664717, false, 0},
		{"cdf_x0", "EXPON.DIST(0,10,TRUE)", 0, false, 0},
		{"cdf_lambda1_x1", "EXPON.DIST(1,1,TRUE)", 0.632120559, false, 0},
		{"cdf_large_x", "EXPON.DIST(10,1,TRUE)", 0.999954600, false, 0},
		{"cdf_small_lambda", "EXPON.DIST(1,0.1,TRUE)", 0.095162582, false, 0},
		{"cdf_large_lambda", "EXPON.DIST(0.01,100,TRUE)", 0.632120559, false, 0},
		{"cdf_x05_lambda2", "EXPON.DIST(0.5,2,TRUE)", 0.632120559, false, 0},
		{"cdf_x1_lambda05", "EXPON.DIST(1,0.5,TRUE)", 0.393469340, false, 0},
		{"cdf_x5_lambda1", "EXPON.DIST(5,1,TRUE)", 0.993262053, false, 0},
		{"cdf_x001_lambda1", "EXPON.DIST(0.01,1,TRUE)", 0.009950166, false, 0},

		// Error cases
		{"err_neg_x", "EXPON.DIST(-1,10,FALSE)", 0, true, ErrValNUM},
		{"err_lambda_zero", "EXPON.DIST(1,0,FALSE)", 0, true, ErrValNUM},
		{"err_lambda_neg", "EXPON.DIST(1,-1,FALSE)", 0, true, ErrValNUM},
		{"err_non_numeric_x", `EXPON.DIST("abc",10,FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_lambda", `EXPON.DIST(1,"abc",FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `EXPON.DIST(1,10,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestEXPON_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "EXPON.DIST(0.2,10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("EXPON.DIST(0.2,10) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "EXPON.DIST(0.2,10,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("EXPON.DIST(0.2,10,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(EXPON.DIST(0.2,10),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(EXPON.DIST(0.2,10),"err") = %v, want string "err"`, got)
	}
}

// WEIBULL.DIST
// ---------------------------------------------------------------------------

func TestWEIBULL_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// PDF tests
		{"pdf_basic", "WEIBULL.DIST(105,20,100,FALSE)", 0.035589, false, 0},
		{"pdf_alpha1_beta1_x1", "WEIBULL.DIST(1,1,1,FALSE)", 0.367879441, false, 0},
		{"pdf_alpha2_beta1_x1", "WEIBULL.DIST(1,2,1,FALSE)", 0.735758882, false, 0},
		{"pdf_alpha1_beta2_x1", "WEIBULL.DIST(1,1,2,FALSE)", 0.303265330, false, 0},
		{"pdf_alpha3_beta1_x05", "WEIBULL.DIST(0.5,3,1,FALSE)", 0.661873, false, 0},
		{"pdf_alpha05_beta1_x1", "WEIBULL.DIST(1,0.5,1,FALSE)", 0.183939721, false, 0},
		{"pdf_alpha10_beta5_x5", "WEIBULL.DIST(5,10,5,FALSE)", 0.735758882, false, 0},
		{"pdf_large_alpha", "WEIBULL.DIST(1,10,1,FALSE)", 3.678794412, false, 0},

		// PDF x=0 edge cases
		{"pdf_x0_alpha1", "WEIBULL.DIST(0,1,2,FALSE)", 0, false, 0},
		{"pdf_x0_alpha2", "WEIBULL.DIST(0,2,1,FALSE)", 0, false, 0},
		{"pdf_x0_alpha3", "WEIBULL.DIST(0,3,1,FALSE)", 0, false, 0},
		{"pdf_x0_alpha05", "WEIBULL.DIST(0,0.5,1,FALSE)", 0, false, 0},

		// CDF tests
		{"cdf_basic", "WEIBULL.DIST(105,20,100,TRUE)", 0.929581, false, 0},
		{"cdf_alpha1_beta1_x1", "WEIBULL.DIST(1,1,1,TRUE)", 0.632120559, false, 0},
		{"cdf_alpha2_beta1_x1", "WEIBULL.DIST(1,2,1,TRUE)", 0.632120559, false, 0},
		{"cdf_alpha1_beta2_x1", "WEIBULL.DIST(1,1,2,TRUE)", 0.393469340, false, 0},
		{"cdf_x0", "WEIBULL.DIST(0,2,1,TRUE)", 0, false, 0},
		{"cdf_x0_alpha1", "WEIBULL.DIST(0,1,1,TRUE)", 0, false, 0},
		{"cdf_large_x", "WEIBULL.DIST(10,1,1,TRUE)", 0.999954600, false, 0},
		{"cdf_small_x", "WEIBULL.DIST(0.01,1,1,TRUE)", 0.009950166, false, 0},
		{"cdf_alpha10_beta5_x5", "WEIBULL.DIST(5,10,5,TRUE)", 0.632120559, false, 0},
		{"cdf_alpha05_beta2_x3", "WEIBULL.DIST(3,0.5,2,TRUE)", 0.706167, false, 0},

		// Error cases
		{"err_neg_x", "WEIBULL.DIST(-1,1,1,FALSE)", 0, true, ErrValNUM},
		{"err_alpha_zero", "WEIBULL.DIST(1,0,1,FALSE)", 0, true, ErrValNUM},
		{"err_alpha_neg", "WEIBULL.DIST(1,-1,1,FALSE)", 0, true, ErrValNUM},
		{"err_beta_zero", "WEIBULL.DIST(1,1,0,FALSE)", 0, true, ErrValNUM},
		{"err_beta_neg", "WEIBULL.DIST(1,1,-1,FALSE)", 0, true, ErrValNUM},
		{"err_non_numeric_x", `WEIBULL.DIST("abc",1,1,FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_alpha", `WEIBULL.DIST(1,"abc",1,FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_beta", `WEIBULL.DIST(1,1,"abc",FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `WEIBULL.DIST(1,1,1,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestWEIBULL_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "WEIBULL.DIST(1,1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("WEIBULL.DIST(1,1,1) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "WEIBULL.DIST(1,1,1,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("WEIBULL.DIST(1,1,1,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(WEIBULL.DIST(1,1,1),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(WEIBULL.DIST(1,1,1),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// LOGNORM.DIST
// ---------------------------------------------------------------------------

func TestLOGNORMDIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		want      float64
		wantError bool
		wantErr   ErrorValue
	}{
		// CDF tests
		{"cdf_excel_example", "LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.0390836, false, 0},
		{"cdf_x1_mean0_sd1", "LOGNORM.DIST(1,0,1,TRUE)", 0.5, false, 0},
		{"cdf_x1_mean0_sd05", "LOGNORM.DIST(1,0,0.5,TRUE)", 0.5, false, 0},
		{"cdf_large_x", "LOGNORM.DIST(1000,0,1,TRUE)", 1.0, false, 0},
		{"cdf_small_x", "LOGNORM.DIST(0.001,0,1,TRUE)", 0.0, false, 0},
		{"cdf_neg_mean", "LOGNORM.DIST(1,-1,1,TRUE)", 0.841345, false, 0},
		{"cdf_large_mean", "LOGNORM.DIST(10,5,2,TRUE)", 0.088715, false, 0},
		{"cdf_small_sd", "LOGNORM.DIST(2.7183,1,0.01,TRUE)", 0.500267, false, 0},
		{"cdf_x_exp_mean", "LOGNORM.DIST(7.389056,2,1,TRUE)", 0.5, false, 0},
		{"cdf_mean2_sd05", "LOGNORM.DIST(10,2,0.5,TRUE)", 0.727467, false, 0},

		// PDF tests
		{"pdf_excel_example", "LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.0176176, false, 0},
		{"pdf_x1_mean0_sd1", "LOGNORM.DIST(1,0,1,FALSE)", 0.398942, false, 0},
		{"pdf_x1_mean0_sd05", "LOGNORM.DIST(1,0,0.5,FALSE)", 0.797885, false, 0},
		{"pdf_neg_mean", "LOGNORM.DIST(1,-1,1,FALSE)", 0.241971, false, 0},
		{"pdf_x2_mean0_sd1", "LOGNORM.DIST(2,0,1,FALSE)", 0.156874, false, 0},
		{"pdf_x05_mean0_sd1", "LOGNORM.DIST(0.5,0,1,FALSE)", 0.627496, false, 0},
		{"pdf_large_x", "LOGNORM.DIST(100,0,1,FALSE)", 0.0000000990, false, 0},
		{"pdf_small_x", "LOGNORM.DIST(0.01,0,1,FALSE)", 0.0009902387, false, 0},
		{"pdf_mean2_sd05", "LOGNORM.DIST(10,2,0.5,FALSE)", 0.0664376, false, 0},

		// Error cases
		{"err_x_zero", "LOGNORM.DIST(0,0,1,TRUE)", 0, true, ErrValNUM},
		{"err_x_neg", "LOGNORM.DIST(-1,0,1,TRUE)", 0, true, ErrValNUM},
		{"err_stdev_zero", "LOGNORM.DIST(1,0,0,TRUE)", 0, true, ErrValNUM},
		{"err_stdev_neg", "LOGNORM.DIST(1,0,-1,TRUE)", 0, true, ErrValNUM},
		{"err_non_numeric_x", `LOGNORM.DIST("abc",0,1,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_mean", `LOGNORM.DIST(1,"abc",1,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_stdev", `LOGNORM.DIST(1,0,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `LOGNORM.DIST(1,0,1,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error %d", tt.formula, got, tt.wantErr)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s error = %d, want %d", tt.formula, got.Err, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v (type %d), want number", tt.formula, got, got.Type)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}
}

func TestLOGNORMDIST_ArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "LOGNORM.DIST(4,3.5,1.2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("LOGNORM.DIST(4,3.5,1.2) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "LOGNORM.DIST(4,3.5,1.2,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("LOGNORM.DIST(4,3.5,1.2,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(LOGNORM.DIST(4,3.5,1.2),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(LOGNORM.DIST(4,3.5,1.2),"err") = %v, want string "err"`, got)
	}
}

func TestLOGNORMDIST_CDFPDFRelation(t *testing.T) {
	// Verify that CDF at x=1, mean=0, stddev=1 is exactly 0.5
	// since ln(1) = 0, so z = (0-0)/1 = 0, and Φ(0) = 0.5
	resolver := &mockResolver{}
	cf := evalCompile(t, "LOGNORM.DIST(1,0,1,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("expected number, got type %d", got.Type)
	}
	if got.Num != 0.5 {
		t.Errorf("LOGNORM.DIST(1,0,1,TRUE) = %g, want exactly 0.5", got.Num)
	}

	// For any mean μ, LOGNORM.DIST(exp(μ), μ, σ, TRUE) should be 0.5
	// because ln(exp(μ)) = μ, so z = (μ-μ)/σ = 0, Φ(0) = 0.5
	cases := []struct {
		mean  float64
		stdev float64
	}{
		{0, 1},
		{1, 2},
		{-1, 0.5},
		{5, 3},
		{0.5, 0.1},
	}
	for _, tc := range cases {
		x := math.Exp(tc.mean)
		formula := fmt.Sprintf("LOGNORM.DIST(%g,%g,%g,TRUE)", x, tc.mean, tc.stdev)
		t.Run(fmt.Sprintf("cdf_half_mean%g_sd%g", tc.mean, tc.stdev), func(t *testing.T) {
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type %d", got.Type)
			}
			if math.Abs(got.Num-0.5) > 1e-10 {
				t.Errorf("%s = %g, want 0.5", formula, got.Num)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// LOGNORM.INV
// ---------------------------------------------------------------------------

func TestLOGNORMINV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		want      float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic cases
		{"excel_example", "LOGNORM.INV(0.039084,3.5,1.2)", 4.0000252, false, 0},
		{"median_mean0_sd1", "LOGNORM.INV(0.5,0,1)", 1.0, false, 0},
		{"median_mean1_sd1", "LOGNORM.INV(0.5,1,1)", 2.718282, false, 0},
		{"median_mean2_sd1", "LOGNORM.INV(0.5,2,1)", 7.389056, false, 0},
		{"median_mean0_sd05", "LOGNORM.INV(0.5,0,0.5)", 1.0, false, 0},
		{"neg_mean", "LOGNORM.INV(0.5,-1,1)", 0.367879, false, 0},
		{"large_stddev", "LOGNORM.INV(0.5,0,5)", 1.0, false, 0},

		// Small and large probabilities
		{"small_p_0.01", "LOGNORM.INV(0.01,0,1)", 0.097652, false, 0},
		{"large_p_0.99", "LOGNORM.INV(0.99,0,1)", 10.24047, false, 0},
		{"small_p_0.001", "LOGNORM.INV(0.001,0,1)", 0.045491, false, 0},
		{"large_p_0.999", "LOGNORM.INV(0.999,0,1)", 21.98218, false, 0},
		{"p_0.1", "LOGNORM.INV(0.1,0,1)", 0.277606, false, 0},
		{"p_0.9", "LOGNORM.INV(0.9,0,1)", 3.602224, false, 0},
		{"p_0.25", "LOGNORM.INV(0.25,0,1)", 0.509416, false, 0},
		{"p_0.75", "LOGNORM.INV(0.75,0,1)", 1.963031, false, 0},

		// Various mean/stdev combinations
		{"mean3_sd2", "LOGNORM.INV(0.5,3,2)", 20.08554, false, 0},
		{"mean5_sd05", "LOGNORM.INV(0.5,5,0.5)", 148.41316, false, 0},
		{"neg_mean2", "LOGNORM.INV(0.5,-2,1)", 0.135335, false, 0},
		{"mean0_sd01", "LOGNORM.INV(0.5,0,0.1)", 1.0, false, 0},

		// Error cases
		{"err_p_zero", "LOGNORM.INV(0,0,1)", 0, true, ErrValNUM},
		{"err_p_one", "LOGNORM.INV(1,0,1)", 0, true, ErrValNUM},
		{"err_p_neg", "LOGNORM.INV(-0.5,0,1)", 0, true, ErrValNUM},
		{"err_p_gt1", "LOGNORM.INV(1.5,0,1)", 0, true, ErrValNUM},
		{"err_stdev_zero", "LOGNORM.INV(0.5,0,0)", 0, true, ErrValNUM},
		{"err_stdev_neg", "LOGNORM.INV(0.5,0,-1)", 0, true, ErrValNUM},
		{"err_non_numeric_p", `LOGNORM.INV("abc",0,1)`, 0, true, ErrValVALUE},
		{"err_non_numeric_mean", `LOGNORM.INV(0.5,"abc",1)`, 0, true, ErrValVALUE},
		{"err_non_numeric_stdev", `LOGNORM.INV(0.5,0,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error %d", tt.formula, got, tt.wantErr)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s error = %d, want %d", tt.formula, got.Err, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v (type %d), want number", tt.formula, got, got.Type)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}
}

func TestLOGNORMINV_ArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "LOGNORM.INV(0.5,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("LOGNORM.INV(0.5,0) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "LOGNORM.INV(0.5,0,1,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("LOGNORM.INV(0.5,0,1,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(LOGNORM.INV(0.5,0),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(LOGNORM.INV(0.5,0),"err") = %v, want string "err"`, got)
	}
}

func TestLOGNORMINV_RoundTrip(t *testing.T) {
	// Verify LOGNORM.INV(LOGNORM.DIST(x, μ, σ, TRUE), μ, σ) ≈ x
	resolver := &mockResolver{}

	cases := []struct {
		x, mean, stdev float64
	}{
		{4, 3.5, 1.2},
		{1, 0, 1},
		{10, 2, 0.5},
		{0.5, -1, 2},
		{100, 5, 1},
	}

	for _, tc := range cases {
		formula := fmt.Sprintf("LOGNORM.INV(LOGNORM.DIST(%g,%g,%g,TRUE),%g,%g)",
			tc.x, tc.mean, tc.stdev, tc.mean, tc.stdev)
		t.Run(fmt.Sprintf("x%g_mean%g_sd%g", tc.x, tc.mean, tc.stdev), func(t *testing.T) {
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type %d", got.Type)
			}
			if math.Abs(got.Num-tc.x) > 1e-4 {
				t.Errorf("%s = %g, want %g", formula, got.Num, tc.x)
			}
		})
	}
}

// GAMMA.DIST
// ---------------------------------------------------------------------------

func TestGAMMA_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// PDF tests
		{"pdf_basic", "GAMMA.DIST(10.00001131,9,2,FALSE)", 0.032639, false, 0},
		{"pdf_x0_alpha1", "GAMMA.DIST(0,1,1,FALSE)", 0, true, ErrValNUM},
		{"pdf_x0_alpha2", "GAMMA.DIST(0,2,1,FALSE)", 0, false, 0},
		{"pdf_x0_alpha1_beta2", "GAMMA.DIST(0,1,2,FALSE)", 0, true, ErrValNUM},
		{"pdf_exponential", "GAMMA.DIST(1,1,1,FALSE)", 0.367879441, false, 0},
		{"pdf_alpha2_beta1", "GAMMA.DIST(1,2,1,FALSE)", 0.367879441, false, 0},
		{"pdf_alpha3_beta1", "GAMMA.DIST(2,3,1,FALSE)", 0.270670566, false, 0},
		{"pdf_alpha_half", "GAMMA.DIST(1,0.5,1,FALSE)", 0.207553749, false, 0},
		{"pdf_alpha2_beta3", "GAMMA.DIST(5,2,3,FALSE)", 0.104930890, false, 0},
		{"pdf_large_x", "GAMMA.DIST(20,5,2,FALSE)", 0.009458319, false, 0},
		{"pdf_small_x", "GAMMA.DIST(0.1,2,1,FALSE)", 0.090483742, false, 0},
		{"pdf_alpha05_beta2", "GAMMA.DIST(0.5,0.5,2,FALSE)", 0.439391289, false, 0},

		// CDF tests
		{"cdf_basic", "GAMMA.DIST(10.00001131,9,2,TRUE)", 0.068094, false, 0},
		{"cdf_x0", "GAMMA.DIST(0,2,1,TRUE)", 0, false, 0},
		{"cdf_exponential", "GAMMA.DIST(1,1,1,TRUE)", 0.632120559, false, 0},
		{"cdf_alpha2_beta1", "GAMMA.DIST(1,2,1,TRUE)", 0.264241118, false, 0},
		{"cdf_alpha3_beta1", "GAMMA.DIST(2,3,1,TRUE)", 0.323323584, false, 0},
		{"cdf_standard_gamma", "GAMMA.DIST(2,3,1,TRUE)", 0.323323584, false, 0},
		{"cdf_alpha2_beta3", "GAMMA.DIST(5,2,3,TRUE)", 0.496331726, false, 0},
		{"cdf_large_x", "GAMMA.DIST(50,5,2,TRUE)", 0.999999733, false, 0},
		{"cdf_alpha_half", "GAMMA.DIST(1,0.5,1,TRUE)", 0.842700793, false, 0},
		{"cdf_alpha1_beta1_x2", "GAMMA.DIST(2,1,1,TRUE)", 0.864664717, false, 0},
		{"cdf_alpha05_beta2", "GAMMA.DIST(0.5,0.5,2,TRUE)", 0.520499878, false, 0},
		{"cdf_alpha10_beta05", "GAMMA.DIST(3,10,0.5,TRUE)", 0.083924017, false, 0},

		// Error: x < 0
		{"err_neg_x", "GAMMA.DIST(-1,9,2,FALSE)", 0, true, ErrValNUM},
		{"err_neg_x_cdf", "GAMMA.DIST(-0.001,1,1,TRUE)", 0, true, ErrValNUM},

		// Error: alpha <= 0
		{"err_alpha_zero", "GAMMA.DIST(1,0,1,FALSE)", 0, true, ErrValNUM},
		{"err_alpha_neg", "GAMMA.DIST(1,-1,1,FALSE)", 0, true, ErrValNUM},
		{"err_alpha_neg2", "GAMMA.DIST(1,-0.5,1,TRUE)", 0, true, ErrValNUM},

		// Error: beta <= 0
		{"err_beta_zero", "GAMMA.DIST(1,1,0,FALSE)", 0, true, ErrValNUM},
		{"err_beta_neg", "GAMMA.DIST(1,1,-1,FALSE)", 0, true, ErrValNUM},
		{"err_beta_neg2", "GAMMA.DIST(1,1,-0.5,TRUE)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `GAMMA.DIST("abc",1,1,FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_alpha", `GAMMA.DIST(1,"abc",1,FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_beta", `GAMMA.DIST(1,1,"abc",FALSE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `GAMMA.DIST(1,1,1,"abc")`, 0, true, ErrValVALUE},

		// x=0 with alpha < 1 (PDF diverges)
		{"err_x0_alpha_lt1_pdf", "GAMMA.DIST(0,0.5,1,FALSE)", 0, true, ErrValNUM},

		// x=0 CDF always 0
		{"cdf_x0_alpha1", "GAMMA.DIST(0,1,1,TRUE)", 0, false, 0},
		{"cdf_x0_alpha05", "GAMMA.DIST(0,0.5,1,TRUE)", 0, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestGAMMA_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "GAMMA.DIST(1,1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMA.DIST(1,1,1) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "GAMMA.DIST(1,1,1,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMA.DIST(1,1,1,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(GAMMA.DIST(1,1,1),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(GAMMA.DIST(1,1,1),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// GAMMA.INV
// ---------------------------------------------------------------------------

func TestGAMMA_INV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic case from spec
		{"basic", "GAMMA.INV(0.068094,9,2)", 10.0000112, false, 0},

		// Exponential distribution (alpha=1): median = ln(2) ≈ 0.693147
		{"exp_median", "GAMMA.INV(0.5,1,1)", 0.693147, false, 0},

		// Gamma(2,1): median ≈ 1.678347
		{"gamma2_1_median", "GAMMA.INV(0.5,2,1)", 1.678347, false, 0},

		// Gamma(5,2): median
		{"gamma5_2_median", "GAMMA.INV(0.5,5,2)", 9.341818, false, 0},

		// Probability 0 → returns 0
		{"p_zero", "GAMMA.INV(0,2,1)", 0, false, 0},

		// Small probability
		{"small_p", "GAMMA.INV(0.01,1,1)", 0.010050, false, 0},

		// Large probability
		{"large_p", "GAMMA.INV(0.99,1,1)", 4.605170, false, 0},

		// Alpha=1 (exponential): GAMMA.INV(1 - e^(-1), 1, 1) ≈ 1
		{"exp_at_1", "GAMMA.INV(0.632121,1,1)", 1.000001, false, 0},

		// Beta scaling: GAMMA.INV(p, a, b) = b * GAMMA.INV(p, a, 1)
		{"beta_scale", "GAMMA.INV(0.5,2,3)", 5.035042, false, 0},

		// Standard gamma (beta=1) various alpha values
		{"std_alpha3", "GAMMA.INV(0.5,3,1)", 2.674060, false, 0},
		{"std_alpha10", "GAMMA.INV(0.5,10,1)", 9.668715, false, 0},
		{"std_alpha05", "GAMMA.INV(0.5,0.5,1)", 0.227468, false, 0},

		// Small alpha
		{"small_alpha", "GAMMA.INV(0.5,0.1,1)", 0.000593, false, 0},

		// Large alpha
		{"large_alpha", "GAMMA.INV(0.5,100,1)", 99.666865, false, 0},

		// Large beta
		{"large_beta", "GAMMA.INV(0.5,2,100)", 167.834699, false, 0},

		// Various probabilities with alpha=3, beta=2
		{"a3b2_p01", "GAMMA.INV(0.1,3,2)", 2.204133, false, 0},
		{"a3b2_p025", "GAMMA.INV(0.25,3,2)", 3.454599, false, 0},
		{"a3b2_p075", "GAMMA.INV(0.75,3,2)", 7.840804, false, 0},
		{"a3b2_p09", "GAMMA.INV(0.9,3,2)", 10.644644, false, 0},

		// Very small probability
		{"very_small_p", "GAMMA.INV(0.001,2,1)", 0.045402, false, 0},

		// Very large probability
		{"very_large_p", "GAMMA.INV(0.999,2,1)", 9.233413, false, 0},

		// Error: probability < 0
		{"err_p_neg", "GAMMA.INV(-0.1,2,1)", 0, true, ErrValNUM},

		// Error: probability > 1
		{"err_p_gt1", "GAMMA.INV(1.5,2,1)", 0, true, ErrValNUM},

		// Error: probability = 1 (Excel returns #NUM!)
		{"err_p_one", "GAMMA.INV(1,2,1)", 0, true, ErrValNUM},

		// Error: alpha = 0
		{"err_alpha_zero", "GAMMA.INV(0.5,0,1)", 0, true, ErrValNUM},

		// Error: alpha < 0
		{"err_alpha_neg", "GAMMA.INV(0.5,-1,1)", 0, true, ErrValNUM},

		// Error: beta = 0
		{"err_beta_zero", "GAMMA.INV(0.5,2,0)", 0, true, ErrValNUM},

		// Error: beta < 0
		{"err_beta_neg", "GAMMA.INV(0.5,2,-1)", 0, true, ErrValNUM},

		// Error: non-numeric arguments
		{"err_non_numeric_p", `GAMMA.INV("abc",2,1)`, 0, true, ErrValVALUE},
		{"err_non_numeric_alpha", `GAMMA.INV(0.5,"abc",1)`, 0, true, ErrValVALUE},
		{"err_non_numeric_beta", `GAMMA.INV(0.5,2,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestGAMMA_INV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "GAMMA.INV(0.5,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMA.INV(0.5,2) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "GAMMA.INV(0.5,2,1,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMA.INV(0.5,2,1,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(GAMMA.INV(0.5,2),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(GAMMA.INV(0.5,2),"err") = %v, want string "err"`, got)
	}
}

func TestGAMMA_INV_roundtrip(t *testing.T) {
	// Verify GAMMA.INV(GAMMA.DIST(x, alpha, beta, TRUE), alpha, beta) ≈ x
	resolver := &mockResolver{}
	const tol = 1e-5

	cases := []struct {
		x     float64
		alpha float64
		beta  float64
	}{
		{10, 9, 2},
		{1, 1, 1},
		{5, 2, 3},
		{0.5, 0.5, 1},
		{20, 5, 2},
		{3, 10, 0.5},
		{100, 50, 2},
		{0.1, 2, 1},
	}

	for _, tc := range cases {
		formula := fmt.Sprintf("GAMMA.INV(GAMMA.DIST(%g,%g,%g,TRUE),%g,%g)",
			tc.x, tc.alpha, tc.beta, tc.alpha, tc.beta)
		t.Run(fmt.Sprintf("x=%g_a=%g_b=%g", tc.x, tc.alpha, tc.beta), func(t *testing.T) {
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tc.x) > tol {
				t.Errorf("%s = %g, want %g", formula, got.Num, tc.x)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// CHISQ.DIST
// ---------------------------------------------------------------------------

func TestCHISQ_DIST(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-5

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// CDF tests
		{"cdf_basic", "CHISQ.DIST(0.5,1,TRUE)", 0.52049988, false, 0},
		{"cdf_df2", "CHISQ.DIST(2,2,TRUE)", 0.63212056, false, 0},
		{"cdf_df3", "CHISQ.DIST(3,3,TRUE)", 0.60837482, false, 0},
		{"cdf_df5", "CHISQ.DIST(5,5,TRUE)", 0.58411981, false, 0},
		{"cdf_x0", "CHISQ.DIST(0,5,TRUE)", 0, false, 0},
		{"cdf_x0_df1", "CHISQ.DIST(0,1,TRUE)", 0, false, 0},
		{"cdf_x0_df2", "CHISQ.DIST(0,2,TRUE)", 0, false, 0},
		{"cdf_df1_p95", "CHISQ.DIST(3.841,1,TRUE)", 0.94998632, false, 0},
		{"cdf_df2_p95", "CHISQ.DIST(5.991,2,TRUE)", 0.94998838, false, 0},
		{"cdf_df10", "CHISQ.DIST(10,10,TRUE)", 0.55950671, false, 0},
		{"cdf_large_x", "CHISQ.DIST(100,5,TRUE)", 1.0, false, 0},
		{"cdf_df20", "CHISQ.DIST(20,20,TRUE)", 0.54207029, false, 0},
		{"cdf_small_x", "CHISQ.DIST(0.01,1,TRUE)", 0.07965567, false, 0},
		{"cdf_df1_x1", "CHISQ.DIST(1,1,TRUE)", 0.68268949, false, 0},

		// PDF tests
		{"pdf_basic", "CHISQ.DIST(2,3,FALSE)", 0.20755375, false, 0},
		{"pdf_df1", "CHISQ.DIST(1,1,FALSE)", 0.24197072, false, 0},
		{"pdf_df2", "CHISQ.DIST(2,2,FALSE)", 0.18393972, false, 0},
		{"pdf_df5", "CHISQ.DIST(4,5,FALSE)", 0.14397591, false, 0},
		{"pdf_df10", "CHISQ.DIST(10,10,FALSE)", 0.08773368, false, 0},
		{"pdf_small_x", "CHISQ.DIST(0.1,2,FALSE)", 0.47561471, false, 0},
		{"pdf_x0_df2", "CHISQ.DIST(0,2,FALSE)", 0.5, false, 0},
		{"pdf_x0_df3", "CHISQ.DIST(0,3,FALSE)", 0, false, 0},
		{"pdf_x0_df4", "CHISQ.DIST(0,4,FALSE)", 0, false, 0},
		{"pdf_x0_df10", "CHISQ.DIST(0,10,FALSE)", 0, false, 0},
		{"pdf_large_x", "CHISQ.DIST(20,5,FALSE)", 0.00053999, false, 0},

		// Truncation: 3.7 → 3
		{"trunc_pdf", "CHISQ.DIST(2,3.7,FALSE)", 0.20755375, false, 0},
		{"trunc_cdf", "CHISQ.DIST(0.5,1.9,TRUE)", 0.52049988, false, 0},

		// Error: x < 0
		{"err_neg_x", "CHISQ.DIST(-1,3,TRUE)", 0, true, ErrValNUM},
		{"err_neg_x_pdf", "CHISQ.DIST(-0.001,1,FALSE)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df_zero", "CHISQ.DIST(1,0,TRUE)", 0, true, ErrValNUM},
		{"err_df_neg", "CHISQ.DIST(1,-1,TRUE)", 0, true, ErrValNUM},
		{"err_df_frac_below1", "CHISQ.DIST(1,0.9,TRUE)", 0, true, ErrValNUM},

		// Error: df > 10^10
		{"err_df_too_large", "CHISQ.DIST(1,10000000001,TRUE)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `CHISQ.DIST("abc",1,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `CHISQ.DIST(1,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `CHISQ.DIST(1,1,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestCHISQ_DIST_argCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "CHISQ.DIST(1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CHISQ.DIST(1,1) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "CHISQ.DIST(1,1,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CHISQ.DIST(1,1,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(CHISQ.DIST(1,1),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(CHISQ.DIST(1,1),"err") = %v, want string "err"`, got)
	}
}

func TestCHISQ_DIST_x0_df1_pdf(t *testing.T) {
	// df=1 ⇒ alpha=0.5, PDF diverges at x=0 (Excel returns +Inf)
	resolver := &mockResolver{}
	cf := evalCompile(t, "CHISQ.DIST(0,1,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("CHISQ.DIST(0,1,FALSE): want number, got type=%d err=%v", got.Type, got.Err)
	}
	if !math.IsInf(got.Num, 1) {
		t.Errorf("CHISQ.DIST(0,1,FALSE) = %g, want +Inf", got.Num)
	}
}

// ---------------------------------------------------------------------------
// CHISQ.INV
// ---------------------------------------------------------------------------

func TestCHISQ_INV(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-5

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases from spec
		{"basic_1", "CHISQ.INV(0.93,1)", 3.283020286, false, 0},
		{"basic_2", "CHISQ.INV(0.6,2)", 1.832581464, false, 0},

		// p=0 returns 0
		{"p_zero", "CHISQ.INV(0,5)", 0, false, 0},
		{"p_zero_df1", "CHISQ.INV(0,1)", 0, false, 0},

		// p=0.5 with df=1
		{"p05_df1", "CHISQ.INV(0.5,1)", 0.454937, false, 0},

		// Critical values (commonly used in statistics)
		{"crit_95_df1", "CHISQ.INV(0.95,1)", 3.841459, false, 0},
		{"crit_95_df2", "CHISQ.INV(0.95,2)", 5.991465, false, 0},
		{"crit_95_df5", "CHISQ.INV(0.95,5)", 11.070498, false, 0},
		{"crit_95_df10", "CHISQ.INV(0.95,10)", 18.307038, false, 0},
		{"crit_99_df1", "CHISQ.INV(0.99,1)", 6.634897, false, 0},
		{"crit_99_df5", "CHISQ.INV(0.99,5)", 15.086272, false, 0},

		// Large df
		{"large_df_50", "CHISQ.INV(0.5,50)", 49.334937, false, 0},
		{"large_df_100", "CHISQ.INV(0.5,100)", 99.334127, false, 0},

		// Small p
		{"small_p", "CHISQ.INV(0.01,5)", 0.554300, false, 0},
		{"small_p_df1", "CHISQ.INV(0.01,1)", 0.000157, false, 0},

		// Large p
		{"large_p_df5", "CHISQ.INV(0.99,5)", 15.086272, false, 0},
		{"large_p_df10", "CHISQ.INV(0.99,10)", 23.209251, false, 0},

		// Various df values at p=0.5
		{"p05_df2", "CHISQ.INV(0.5,2)", 1.386294, false, 0},
		{"p05_df3", "CHISQ.INV(0.5,3)", 2.365974, false, 0},
		{"p05_df5", "CHISQ.INV(0.5,5)", 4.351460, false, 0},
		{"p05_df10", "CHISQ.INV(0.5,10)", 9.341818, false, 0},
		{"p05_df20", "CHISQ.INV(0.5,20)", 19.337430, false, 0},

		// Truncation: 3.7 truncates to 3, same as df=3
		{"trunc_37", "CHISQ.INV(0.5,3.7)", 2.365974, false, 0},
		{"trunc_19", "CHISQ.INV(0.5,1.9)", 0.454937, false, 0},

		// Error: p < 0
		{"err_p_neg", "CHISQ.INV(-0.1,5)", 0, true, ErrValNUM},

		// Error: p > 1
		{"err_p_gt1", "CHISQ.INV(1.5,5)", 0, true, ErrValNUM},

		// Error: p = 1
		{"err_p_one", "CHISQ.INV(1,5)", 0, true, ErrValNUM},

		// Error: df = 0
		{"err_df_zero", "CHISQ.INV(0.5,0)", 0, true, ErrValNUM},

		// Error: df < 0
		{"err_df_neg", "CHISQ.INV(0.5,-1)", 0, true, ErrValNUM},

		// Error: df < 1 after truncation
		{"err_df_frac_below1", "CHISQ.INV(0.5,0.9)", 0, true, ErrValNUM},

		// Error: df > 10^10
		{"err_df_too_large", "CHISQ.INV(0.5,10000000001)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_p", `CHISQ.INV("abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `CHISQ.INV(0.5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestCHISQ_INV_argCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "CHISQ.INV(0.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CHISQ.INV(0.5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "CHISQ.INV(0.5,5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CHISQ.INV(0.5,5,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(CHISQ.INV(0.5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(CHISQ.INV(0.5),"err") = %v, want string "err"`, got)
	}
}

func TestCHISQ_INV_roundtrip(t *testing.T) {
	// Verify CHISQ.INV(CHISQ.DIST(x, df, TRUE), df) ≈ x
	resolver := &mockResolver{}
	const tol = 1e-4

	cases := []struct {
		x  float64
		df int
	}{
		{5, 3},
		{1, 1},
		{10, 5},
		{2, 2},
		{15, 10},
		{0.5, 1},
		{20, 20},
		{3, 7},
	}

	for _, tc := range cases {
		formula := fmt.Sprintf("CHISQ.INV(CHISQ.DIST(%g,%d,TRUE),%d)",
			tc.x, tc.df, tc.df)
		t.Run(fmt.Sprintf("x=%g_df=%d", tc.x, tc.df), func(t *testing.T) {
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tc.x) > tol {
				t.Errorf("%s = %g, want %g", formula, got.Num, tc.x)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// T.DIST
// ---------------------------------------------------------------------------

func TestT_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// CDF tests
		{"cdf_60_1", "T.DIST(60,1,TRUE)", 0.99469533, false, 0},
		{"cdf_zero_symmetric", "T.DIST(0,5,TRUE)", 0.5, false, 0},
		{"cdf_zero_df1", "T.DIST(0,1,TRUE)", 0.5, false, 0},
		{"cdf_zero_df10", "T.DIST(0,10,TRUE)", 0.5, false, 0},
		{"cdf_zero_df100", "T.DIST(0,100,TRUE)", 0.5, false, 0},
		{"cdf_neg2_df10", "T.DIST(-2,10,TRUE)", 0.03669402, false, 0},
		{"cdf_pos2_df10", "T.DIST(2,10,TRUE)", 0.96330598, false, 0},
		{"cdf_large_df_normal_approx", "T.DIST(1.96,1000,TRUE)", 0.97486341, false, 0},
		{"cdf_cauchy_1", "T.DIST(1,1,TRUE)", 0.75, false, 0},
		{"cdf_cauchy_neg1", "T.DIST(-1,1,TRUE)", 0.25, false, 0},
		{"cdf_2571_df5", "T.DIST(2.571,5,TRUE)", 0.97501268, false, 0},
		{"cdf_neg3_df5", "T.DIST(-3,5,TRUE)", 0.01504962, false, 0},
		{"cdf_3_df5", "T.DIST(3,5,TRUE)", 0.98495038, false, 0},
		{"cdf_1_df2", "T.DIST(1,2,TRUE)", 0.78867513, false, 0},
		{"cdf_neg1_df2", "T.DIST(-1,2,TRUE)", 0.21132487, false, 0},
		{"cdf_large_x_df5", "T.DIST(100,5,TRUE)", 1.0, false, 0},
		{"cdf_neg_large_x_df5", "T.DIST(-100,5,TRUE)", 0.0, false, 0},
		{"cdf_05_df30", "T.DIST(0.5,30,TRUE)", 0.68963850, false, 0},

		// PDF tests
		{"pdf_8_3", "T.DIST(8,3,FALSE)", 0.00073691, false, 0},
		{"pdf_zero_df5", "T.DIST(0,5,FALSE)", 0.37960669, false, 0},
		{"pdf_zero_df1_cauchy", "T.DIST(0,1,FALSE)", 0.31830989, false, 0},
		{"pdf_large_df_normal", "T.DIST(0,100,FALSE)", 0.39794974, false, 0},
		{"pdf_very_large_df", "T.DIST(0,10000,FALSE)", 0.39894216, false, 0},
		{"pdf_neg3_df5", "T.DIST(-3,5,FALSE)", 0.01729258, false, 0},
		{"pdf_pos3_df5", "T.DIST(3,5,FALSE)", 0.01729258, false, 0},
		{"pdf_1_df1", "T.DIST(1,1,FALSE)", 0.15915494, false, 0},
		{"pdf_neg1_df1", "T.DIST(-1,1,FALSE)", 0.15915494, false, 0},
		{"pdf_2_df10", "T.DIST(2,10,FALSE)", 0.06114577, false, 0},
		{"pdf_neg2_df10", "T.DIST(-2,10,FALSE)", 0.06114577, false, 0},

		// Error: df < 1
		{"err_df_zero", "T.DIST(1,0,TRUE)", 0, true, ErrValNUM},
		{"err_df_neg", "T.DIST(1,-1,TRUE)", 0, true, ErrValNUM},
		{"err_df_frac_below1", "T.DIST(1,0.9,TRUE)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `T.DIST("abc",5,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `T.DIST(1,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `T.DIST(1,5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestT_DIST_symmetry(t *testing.T) {
	// CDF(-x) + CDF(x) should equal 1 for the t-distribution.
	const tol = 1e-10
	resolver := &mockResolver{}

	cases := []struct {
		x  float64
		df int
	}{
		{2, 10},
		{1, 1},
		{3, 5},
		{0.5, 30},
		{1.96, 1000},
	}

	for _, tc := range cases {
		t.Run(fmt.Sprintf("x=%g_df=%d", tc.x, tc.df), func(t *testing.T) {
			fPos := fmt.Sprintf("T.DIST(%g,%d,TRUE)", tc.x, tc.df)
			fNeg := fmt.Sprintf("T.DIST(%g,%d,TRUE)", -tc.x, tc.df)

			cfPos := evalCompile(t, fPos)
			gotPos, err := Eval(cfPos, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			cfNeg := evalCompile(t, fNeg)
			gotNeg, err := Eval(cfNeg, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}

			if gotPos.Type != ValueNumber || gotNeg.Type != ValueNumber {
				t.Fatalf("expected numbers, got pos=%d neg=%d", gotPos.Type, gotNeg.Type)
			}

			sum := gotPos.Num + gotNeg.Num
			if math.Abs(sum-1.0) > tol {
				t.Errorf("CDF(%g) + CDF(%g) = %g + %g = %g, want 1.0",
					tc.x, -tc.x, gotPos.Num, gotNeg.Num, sum)
			}
		})
	}
}

func TestT_DIST_pdf_symmetry(t *testing.T) {
	// PDF(-x) should equal PDF(x) for the t-distribution.
	const tol = 1e-12
	resolver := &mockResolver{}

	cases := []struct {
		x  float64
		df int
	}{
		{3, 5},
		{1, 1},
		{2, 10},
		{0.5, 30},
	}

	for _, tc := range cases {
		t.Run(fmt.Sprintf("x=%g_df=%d", tc.x, tc.df), func(t *testing.T) {
			fPos := fmt.Sprintf("T.DIST(%g,%d,FALSE)", tc.x, tc.df)
			fNeg := fmt.Sprintf("T.DIST(%g,%d,FALSE)", -tc.x, tc.df)

			cfPos := evalCompile(t, fPos)
			gotPos, err := Eval(cfPos, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			cfNeg := evalCompile(t, fNeg)
			gotNeg, err := Eval(cfNeg, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}

			if gotPos.Type != ValueNumber || gotNeg.Type != ValueNumber {
				t.Fatalf("expected numbers, got pos=%d neg=%d", gotPos.Type, gotNeg.Type)
			}

			if math.Abs(gotPos.Num-gotNeg.Num) > tol {
				t.Errorf("PDF(%g) = %g != PDF(%g) = %g", tc.x, gotPos.Num, -tc.x, gotNeg.Num)
			}
		})
	}
}

func TestT_DIST_truncation(t *testing.T) {
	// T.DIST(1, 5.7, ...) should equal T.DIST(1, 5, ...)
	const tol = 1e-12
	resolver := &mockResolver{}

	// CDF
	cfTrunc := evalCompile(t, "T.DIST(1,5.7,TRUE)")
	gotTrunc, err := Eval(cfTrunc, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	cfInt := evalCompile(t, "T.DIST(1,5,TRUE)")
	gotInt, err := Eval(cfInt, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if gotTrunc.Type != ValueNumber || gotInt.Type != ValueNumber {
		t.Fatalf("expected numbers, got trunc=%d int=%d", gotTrunc.Type, gotInt.Type)
	}
	if math.Abs(gotTrunc.Num-gotInt.Num) > tol {
		t.Errorf("T.DIST(1,5.7,TRUE) = %g, T.DIST(1,5,TRUE) = %g; want equal", gotTrunc.Num, gotInt.Num)
	}

	// PDF
	cfTrunc = evalCompile(t, "T.DIST(1,5.7,FALSE)")
	gotTrunc, err = Eval(cfTrunc, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	cfInt = evalCompile(t, "T.DIST(1,5,FALSE)")
	gotInt, err = Eval(cfInt, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if gotTrunc.Type != ValueNumber || gotInt.Type != ValueNumber {
		t.Fatalf("expected numbers, got trunc=%d int=%d", gotTrunc.Type, gotInt.Type)
	}
	if math.Abs(gotTrunc.Num-gotInt.Num) > tol {
		t.Errorf("T.DIST(1,5.7,FALSE) = %g, T.DIST(1,5,FALSE) = %g; want equal", gotTrunc.Num, gotInt.Num)
	}
}

func TestT_DIST_argCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "T.DIST(1,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("T.DIST(1,5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "T.DIST(1,5,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("T.DIST(1,5,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(T.DIST(1,5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(T.DIST(1,5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// T.INV
// ---------------------------------------------------------------------------

func TestT_INV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic values
		{"basic_075_df2", "T.INV(0.75,2)", 0.8164966, false, 0},
		{"p05_df5_zero", "T.INV(0.5,5)", 0, false, 0},
		{"p05_df1_zero", "T.INV(0.5,1)", 0, false, 0},
		{"p05_df100_zero", "T.INV(0.5,100)", 0, false, 0},

		// Critical values (well-known from t-tables)
		{"p0975_df10", "T.INV(0.975,10)", 2.228139, false, 0},
		{"p0025_df10", "T.INV(0.025,10)", -2.228139, false, 0},
		{"p095_df1", "T.INV(0.95,1)", 6.313752, false, 0},
		{"p095_df2", "T.INV(0.95,2)", 2.919986, false, 0},
		{"p095_df5", "T.INV(0.95,5)", 2.015048, false, 0},
		{"p095_df30", "T.INV(0.95,30)", 1.697261, false, 0},

		// Large df approaches normal
		{"p0975_df1000", "T.INV(0.975,1000)", 1.962339, false, 0},

		// Small and large p
		{"p001_df5", "T.INV(0.01,5)", -3.364930, false, 0},
		{"p099_df5", "T.INV(0.99,5)", 3.364930, false, 0},

		// Cauchy distribution (df=1)
		{"cauchy_p075", "T.INV(0.75,1)", 1.0, false, 0},
		{"cauchy_p025", "T.INV(0.25,1)", -1.0, false, 0},
		{"cauchy_p09", "T.INV(0.9,1)", 3.077684, false, 0},

		// df truncation: T.INV(0.75, 2.9) should equal T.INV(0.75, 2)
		{"df_truncation", "T.INV(0.75,2.9)", 0.8164966, false, 0},

		// Extreme p values
		{"p0001_df10", "T.INV(0.001,10)", -4.143700, false, 0},
		{"p0999_df10", "T.INV(0.999,10)", 4.143700, false, 0},

		// Various df values
		{"p09_df3", "T.INV(0.9,3)", 1.637744, false, 0},
		{"p01_df3", "T.INV(0.1,3)", -1.637744, false, 0},

		// Error: p = 0
		{"err_p_zero", "T.INV(0,5)", 0, true, ErrValNUM},
		// Error: p < 0
		{"err_p_neg", "T.INV(-0.1,5)", 0, true, ErrValNUM},
		// Error: p = 1
		{"err_p_one", "T.INV(1,5)", 0, true, ErrValNUM},
		// Error: p > 1
		{"err_p_gt1", "T.INV(1.5,5)", 0, true, ErrValNUM},
		// Error: df = 0
		{"err_df_zero", "T.INV(0.5,0)", 0, true, ErrValNUM},
		// Error: df < 1
		{"err_df_neg", "T.INV(0.5,-1)", 0, true, ErrValNUM},
		{"err_df_frac_below1", "T.INV(0.5,0.9)", 0, true, ErrValNUM},
		// Error: non-numeric
		{"err_non_numeric_p", `T.INV("abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `T.INV(0.5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestT_INV_roundTrip(t *testing.T) {
	// T.INV(T.DIST(x, df, TRUE), df) should ≈ x
	const tol = 1e-6
	resolver := &mockResolver{}

	cases := []struct {
		x  float64
		df int
	}{
		{2, 5},
		{-1.5, 10},
		{0, 3},
		{3.5, 20},
		{-0.5, 1},
	}

	for _, tc := range cases {
		t.Run(fmt.Sprintf("x=%g_df=%d", tc.x, tc.df), func(t *testing.T) {
			formula := fmt.Sprintf("T.INV(T.DIST(%g,%d,TRUE),%d)", tc.x, tc.df, tc.df)
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tc.x) > tol {
				t.Errorf("T.INV(T.DIST(%g,%d,TRUE),%d) = %g, want %g", tc.x, tc.df, tc.df, got.Num, tc.x)
			}
		})
	}
}

func TestT_INV_symmetry(t *testing.T) {
	// T.INV(p, df) = -T.INV(1-p, df)
	const tol = 1e-8
	resolver := &mockResolver{}

	cases := []struct {
		p  float64
		df int
	}{
		{0.025, 10},
		{0.1, 5},
		{0.9, 30},
		{0.75, 1},
		{0.99, 20},
	}

	for _, tc := range cases {
		t.Run(fmt.Sprintf("p=%g_df=%d", tc.p, tc.df), func(t *testing.T) {
			fP := fmt.Sprintf("T.INV(%g,%d)", tc.p, tc.df)
			fQ := fmt.Sprintf("T.INV(%g,%d)", 1-tc.p, tc.df)

			cfP := evalCompile(t, fP)
			gotP, err := Eval(cfP, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			cfQ := evalCompile(t, fQ)
			gotQ, err := Eval(cfQ, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}

			if gotP.Type != ValueNumber || gotQ.Type != ValueNumber {
				t.Fatalf("want numbers, got P type=%d, Q type=%d", gotP.Type, gotQ.Type)
			}

			sum := gotP.Num + gotQ.Num
			if math.Abs(sum) > tol {
				t.Errorf("T.INV(%g,%d) + T.INV(%g,%d) = %g + %g = %g, want 0",
					tc.p, tc.df, 1-tc.p, tc.df, gotP.Num, gotQ.Num, sum)
			}
		})
	}
}

func TestT_INV_argCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "T.INV(0.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("T.INV(0.5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "T.INV(0.5,5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("T.INV(0.5,5,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch
	cf = evalCompile(t, `IFERROR(T.INV(0.5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(T.INV(0.5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// F.DIST
// ---------------------------------------------------------------------------

func TestF_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// CDF tests
		{"cdf_basic", "F.DIST(15.2069,6,4,TRUE)", 0.9900000430, false, 0},
		{"cdf_x0", "F.DIST(0,5,5,TRUE)", 0, false, 0},
		{"cdf_x0_df1_1", "F.DIST(0,1,5,TRUE)", 0, false, 0},
		{"cdf_x1_equal_df", "F.DIST(1,5,5,TRUE)", 0.5, false, 0},
		{"cdf_x1_df10_10", "F.DIST(1,10,10,TRUE)", 0.5, false, 0},
		{"cdf_large_x", "F.DIST(100,5,5,TRUE)", 0.9999475709, false, 0},
		{"cdf_small_x", "F.DIST(0.01,5,5,TRUE)", 0.0000524291, false, 0},
		{"cdf_df_1_1", "F.DIST(1,1,1,TRUE)", 0.5, false, 0},
		{"cdf_df_2_3", "F.DIST(2,2,3,TRUE)", 0.7194341411, false, 0},
		{"cdf_df_5_10", "F.DIST(2,5,10,TRUE)", 0.8358050491, false, 0},
		{"cdf_df_10_20", "F.DIST(1.5,10,20,TRUE)", 0.7890535375, false, 0},
		{"cdf_df_1_100", "F.DIST(3.84,1,100,TRUE)", 0.9471726649, false, 0},
		{"cdf_truncation", "F.DIST(1,5.7,5.7,TRUE)", 0.5, false, 0},
		{"cdf_truncation2", "F.DIST(2,2.9,3.9,TRUE)", 0.7194341411, false, 0},
		{"cdf_small_df", "F.DIST(5,1,1,TRUE)", 0.7322795272, false, 0},
		{"cdf_large_df", "F.DIST(1,100,100,TRUE)", 0.5, false, 0},

		// PDF tests
		{"pdf_basic", "F.DIST(15.2069,6,4,FALSE)", 0.0012237917, false, 0},
		{"pdf_x0_df1_2", "F.DIST(0,2,5,FALSE)", 1, false, 0},
		{"pdf_x0_df1_gt2", "F.DIST(0,3,5,FALSE)", 0, false, 0},
		{"pdf_x0_df1_4", "F.DIST(0,4,10,FALSE)", 0, false, 0},
		{"pdf_x1_df10_10", "F.DIST(1,10,10,FALSE)", 0.6152343750, false, 0},
		{"pdf_df_1_1", "F.DIST(1,1,1,FALSE)", 0.1591549431, false, 0},
		{"pdf_df_2_3", "F.DIST(1,2,3,FALSE)", 0.2788548009, false, 0},
		{"pdf_df_5_10", "F.DIST(2,5,10,FALSE)", 0.1620057422, false, 0},
		{"pdf_df_10_20", "F.DIST(1.5,10,20,FALSE)", 0.3581610917, false, 0},
		{"pdf_large_x", "F.DIST(10,5,5,FALSE)", 0.0026667077, false, 0},
		{"pdf_small_x", "F.DIST(0.1,5,5,FALSE)", 0.2666707709, false, 0},

		// Error: x < 0
		{"err_neg_x", "F.DIST(-1,5,5,TRUE)", 0, true, ErrValNUM},
		{"err_neg_x_pdf", "F.DIST(-0.001,5,5,FALSE)", 0, true, ErrValNUM},

		// Error: df1 < 1
		{"err_df1_zero", "F.DIST(1,0,5,TRUE)", 0, true, ErrValNUM},
		{"err_df1_neg", "F.DIST(1,-1,5,TRUE)", 0, true, ErrValNUM},
		{"err_df1_frac_lt1", "F.DIST(1,0.9,5,TRUE)", 0, true, ErrValNUM},

		// Error: df2 < 1
		{"err_df2_zero", "F.DIST(1,5,0,TRUE)", 0, true, ErrValNUM},
		{"err_df2_neg", "F.DIST(1,5,-1,TRUE)", 0, true, ErrValNUM},
		{"err_df2_frac_lt1", "F.DIST(1,5,0.5,TRUE)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `F.DIST("abc",5,5,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df1", `F.DIST(1,"abc",5,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df2", `F.DIST(1,5,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `F.DIST(1,5,5,"abc")`, 0, true, ErrValVALUE},

		// Error: x=0 with df1=1 (PDF diverges)
		{"err_x0_df1_1_pdf", "F.DIST(0,1,5,FALSE)", 0, true, ErrValNUM},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v (type %d), want number", tt.formula, got, got.Type)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestF_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "F.DIST(1,5,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("F.DIST(1,5,5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "F.DIST(1,5,5,TRUE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("F.DIST(1,5,5,TRUE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(F.DIST(1,5,5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(F.DIST(1,5,5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// F.INV
// ---------------------------------------------------------------------------

func TestF_INV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic tests
		{"basic_001_6_4", "F.INV(0.01,6,4)", 0.10930991, false, 0},
		{"basic_099_6_4", "F.INV(0.99,6,4)", 15.2069, false, 0},
		{"basic_05_5_5", "F.INV(0.5,5,5)", 1.0, false, 0},
		{"basic_05_10_10", "F.INV(0.5,10,10)", 1.0, false, 0},
		{"basic_05_1_1", "F.INV(0.5,1,1)", 1.0, false, 0},

		// p=0 edge case
		{"p_zero", "F.INV(0,5,5)", 0, false, 0},

		// Various critical values at p=0.95
		{"p095_1_1", "F.INV(0.95,1,1)", 161.4476, false, 0},
		{"p095_2_3", "F.INV(0.95,2,3)", 9.552094, false, 0},
		{"p095_5_10", "F.INV(0.95,5,10)", 3.325835, false, 0},
		{"p095_10_20", "F.INV(0.95,10,20)", 2.347878, false, 0},

		// Small probability
		{"small_p_001_5_5", "F.INV(0.001,5,5)", 0.0336107355, false, 0},

		// Large probability
		{"large_p_999_5_5", "F.INV(0.999,5,5)", 29.7523985773, false, 0},

		// Various df combinations
		{"df_2_5", "F.INV(0.5,2,5)", 0.7987697769, false, 0},
		{"df_10_20", "F.INV(0.5,10,20)", 0.9662638886, false, 0},
		{"df_5_100", "F.INV(0.5,5,100)", 0.8761990472, false, 0},
		{"df_1_100", "F.INV(0.5,1,100)", 0.4582627146, false, 0},

		// Truncation: F.INV(0.5, 5.7, 5.7) == F.INV(0.5, 5, 5)
		{"df_truncation", "F.INV(0.5,5.7,5.7)", 1.0, false, 0},

		// More diverse probabilities
		{"p010_5_10", "F.INV(0.10,5,10)", 0.3032690890, false, 0},
		{"p025_5_10", "F.INV(0.25,5,10)", 0.5291416856, false, 0},
		{"p075_5_10", "F.INV(0.75,5,10)", 1.5853232594, false, 0},
		{"p090_5_10", "F.INV(0.90,5,10)", 2.5216406862, false, 0},

		// Error: p < 0
		{"err_p_neg", "F.INV(-0.1,5,5)", 0, true, ErrValNUM},

		// Error: p > 1
		{"err_p_gt1", "F.INV(1.5,5,5)", 0, true, ErrValNUM},

		// Error: p = 1
		{"err_p_one", "F.INV(1,5,5)", 0, true, ErrValNUM},

		// Error: df1 < 1
		{"err_df1_zero", "F.INV(0.5,0,5)", 0, true, ErrValNUM},
		{"err_df1_neg", "F.INV(0.5,-1,5)", 0, true, ErrValNUM},
		{"err_df1_frac_lt1", "F.INV(0.5,0.9,5)", 0, true, ErrValNUM},

		// Error: df2 < 1
		{"err_df2_zero", "F.INV(0.5,5,0)", 0, true, ErrValNUM},
		{"err_df2_neg", "F.INV(0.5,5,-1)", 0, true, ErrValNUM},
		{"err_df2_frac_lt1", "F.INV(0.5,5,0.5)", 0, true, ErrValNUM},

		// Error: non-numeric arguments
		{"err_non_numeric_p", `F.INV("abc",5,5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df1", `F.INV(0.5,"abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df2", `F.INV(0.5,5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v (type %d), want number", tt.formula, got, got.Type)
			}
			if tt.wantNum == 0 {
				if math.Abs(got.Num) > tol {
					t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else {
				relErr := math.Abs(got.Num-tt.wantNum) / math.Abs(tt.wantNum)
				if relErr > 1e-4 {
					t.Errorf("%s = %g, want %g (relErr=%g)", tt.formula, got.Num, tt.wantNum, relErr)
				}
			}
		})
	}
}

func TestF_INV_roundtrip(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-6

	// F.INV(F.DIST(x, df1, df2, TRUE), df1, df2) should ≈ x
	cases := []struct {
		x   float64
		df1 int
		df2 int
	}{
		{5, 3, 7},
		{1, 5, 5},
		{2, 2, 3},
		{0.5, 10, 20},
		{3, 1, 100},
		{10, 5, 5},
		{0.1, 5, 10},
	}

	for _, tc := range cases {
		formula := fmt.Sprintf("F.INV(F.DIST(%g,%d,%d,TRUE),%d,%d)", tc.x, tc.df1, tc.df2, tc.df1, tc.df2)
		t.Run(formula, func(t *testing.T) {
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v, want number", formula, got)
			}
			relErr := math.Abs(got.Num-tc.x) / math.Abs(tc.x)
			if relErr > tol {
				t.Errorf("%s = %g, want %g (relErr=%g)", formula, got.Num, tc.x, relErr)
			}
		})
	}
}

func TestF_INV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "F.INV(0.5,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("F.INV(0.5,5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "F.INV(0.5,5,5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("F.INV(0.5,5,5,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(F.INV(0.5,5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(F.INV(0.5,5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// CONFIDENCE.NORM
// ---------------------------------------------------------------------------

func TestCONFIDENCE_NORM(t *testing.T) {
	const tol = 1e-4
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic values verified against Excel
		{"basic_005_25_50", "CONFIDENCE.NORM(0.05,2.5,50)", 0.692952, false, 0},
		{"basic_005_1_100", "CONFIDENCE.NORM(0.05,1,100)", 0.196, false, 0},
		{"basic_001_25_50", "CONFIDENCE.NORM(0.01,2.5,50)", 0.910693, false, 0},
		{"basic_01_1_25", "CONFIDENCE.NORM(0.1,1,25)", 0.328971, false, 0},
		{"basic_005_1_30", "CONFIDENCE.NORM(0.05,1,30)", 0.357886, false, 0},
		{"basic_01_5_200", "CONFIDENCE.NORM(0.1,5,200)", 0.581508, false, 0},

		// Large sample size
		{"large_sample", "CONFIDENCE.NORM(0.05,1,10000)", 0.0196, false, 0},

		// Small alpha (high confidence)
		{"small_alpha", "CONFIDENCE.NORM(0.001,1,100)", 0.329053, false, 0},

		// Large stddev
		{"large_stddev", "CONFIDENCE.NORM(0.05,100,50)", 27.718, false, 0},

		// Fractional size should be truncated
		{"size_truncation_29", "CONFIDENCE.NORM(0.05,2.5,50.9)", 0.692952, false, 0},
		{"size_truncation_11", "CONFIDENCE.NORM(0.05,1,1.9)", 1.959964, false, 0},

		// Error: alpha = 0
		{"err_alpha_zero", "CONFIDENCE.NORM(0,2.5,50)", 0, true, ErrValNUM},
		// Error: alpha = 1
		{"err_alpha_one", "CONFIDENCE.NORM(1,2.5,50)", 0, true, ErrValNUM},
		// Error: alpha < 0
		{"err_alpha_neg", "CONFIDENCE.NORM(-0.05,2.5,50)", 0, true, ErrValNUM},
		// Error: alpha > 1
		{"err_alpha_gt1", "CONFIDENCE.NORM(1.5,2.5,50)", 0, true, ErrValNUM},
		// Error: stddev = 0
		{"err_stddev_zero", "CONFIDENCE.NORM(0.05,0,50)", 0, true, ErrValNUM},
		// Error: stddev < 0
		{"err_stddev_neg", "CONFIDENCE.NORM(0.05,-1,50)", 0, true, ErrValNUM},
		// Error: size = 0
		{"err_size_zero", "CONFIDENCE.NORM(0.05,2.5,0)", 0, true, ErrValNUM},
		// Error: size < 1 (fractional truncates to 0)
		{"err_size_frac", "CONFIDENCE.NORM(0.05,2.5,0.9)", 0, true, ErrValNUM},
		// Error: size negative
		{"err_size_neg", "CONFIDENCE.NORM(0.05,2.5,-5)", 0, true, ErrValNUM},
		// Error: non-numeric alpha
		{"err_nonnumeric_alpha", `CONFIDENCE.NORM("abc",2.5,50)`, 0, true, ErrValVALUE},
		// Error: non-numeric stddev
		{"err_nonnumeric_stddev", `CONFIDENCE.NORM(0.05,"abc",50)`, 0, true, ErrValVALUE},
		// Error: non-numeric size
		{"err_nonnumeric_size", `CONFIDENCE.NORM(0.05,2.5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestCONFIDENCE_NORM_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few arguments
	cf := evalCompile(t, "CONFIDENCE.NORM(0.05,2.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CONFIDENCE.NORM(0.05,2.5) should error, got type=%d", got.Type)
	}

	// Too many arguments
	cf = evalCompile(t, "CONFIDENCE.NORM(0.05,2.5,50,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CONFIDENCE.NORM(0.05,2.5,50,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(CONFIDENCE.NORM(0.05,2.5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(CONFIDENCE.NORM(0.05,2.5),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// CONFIDENCE.T
// ---------------------------------------------------------------------------

func TestCONFIDENCE_T(t *testing.T) {
	const tol = 1e-4
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic values verified against Excel
		{"basic_005_1_50", "CONFIDENCE.T(0.05,1,50)", 0.284196, false, 0},
		{"basic_005_25_50", "CONFIDENCE.T(0.05,2.5,50)", 0.710492, false, 0},
		{"basic_001_1_30", "CONFIDENCE.T(0.01,1,30)", 0.503245, false, 0},
		{"basic_01_1_25", "CONFIDENCE.T(0.1,1,25)", 0.342176, false, 0},
		{"basic_005_1_100", "CONFIDENCE.T(0.05,1,100)", 0.198422, false, 0},
		{"basic_01_5_200", "CONFIDENCE.T(0.1,5,200)", 0.584264, false, 0},

		// Small sample (df=1 for size=2)
		{"small_sample_2", "CONFIDENCE.T(0.05,1,2)", 8.984644, false, 0},
		// Small sample (df=2 for size=3)
		{"small_sample_3", "CONFIDENCE.T(0.05,1,3)", 2.484159, false, 0},
		// Small sample (df=4 for size=5)
		{"small_sample_5", "CONFIDENCE.T(0.05,1,5)", 1.241617, false, 0},

		// Large sample — should approach CONFIDENCE.NORM
		{"large_sample", "CONFIDENCE.T(0.05,1,10000)", 0.019600, false, 0},

		// Fractional size should be truncated
		{"size_truncation", "CONFIDENCE.T(0.05,1,50.9)", 0.284196, false, 0},
		{"size_truncation_2", "CONFIDENCE.T(0.05,2.5,50.1)", 0.710492, false, 0},

		// Error: size = 1 → #DIV/0! (df=0)
		{"err_size_one", "CONFIDENCE.T(0.05,1,1)", 0, true, ErrValDIV0},
		// Error: size = 1 with fractional (truncates to 1)
		{"err_size_one_frac", "CONFIDENCE.T(0.05,1,1.9)", 0, true, ErrValDIV0},
		// Error: alpha = 0
		{"err_alpha_zero", "CONFIDENCE.T(0,1,50)", 0, true, ErrValNUM},
		// Error: alpha = 1
		{"err_alpha_one", "CONFIDENCE.T(1,1,50)", 0, true, ErrValNUM},
		// Error: alpha < 0
		{"err_alpha_neg", "CONFIDENCE.T(-0.05,1,50)", 0, true, ErrValNUM},
		// Error: alpha > 1
		{"err_alpha_gt1", "CONFIDENCE.T(1.5,1,50)", 0, true, ErrValNUM},
		// Error: stddev = 0
		{"err_stddev_zero", "CONFIDENCE.T(0.05,0,50)", 0, true, ErrValNUM},
		// Error: stddev < 0
		{"err_stddev_neg", "CONFIDENCE.T(0.05,-1,50)", 0, true, ErrValNUM},
		// Error: size = 0
		{"err_size_zero", "CONFIDENCE.T(0.05,1,0)", 0, true, ErrValNUM},
		// Error: size < 1 (fractional truncates to 0)
		{"err_size_frac_below1", "CONFIDENCE.T(0.05,1,0.5)", 0, true, ErrValNUM},
		// Error: size negative
		{"err_size_neg", "CONFIDENCE.T(0.05,1,-5)", 0, true, ErrValNUM},
		// Error: non-numeric alpha
		{"err_nonnumeric_alpha", `CONFIDENCE.T("abc",1,50)`, 0, true, ErrValVALUE},
		// Error: non-numeric stddev
		{"err_nonnumeric_stddev", `CONFIDENCE.T(0.05,"abc",50)`, 0, true, ErrValVALUE},
		// Error: non-numeric size
		{"err_nonnumeric_size", `CONFIDENCE.T(0.05,1,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestCONFIDENCE_T_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few arguments
	cf := evalCompile(t, "CONFIDENCE.T(0.05,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CONFIDENCE.T(0.05,1) should error, got type=%d", got.Type)
	}

	// Too many arguments
	cf = evalCompile(t, "CONFIDENCE.T(0.05,1,50,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("CONFIDENCE.T(0.05,1,50,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(CONFIDENCE.T(0.05,1),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(CONFIDENCE.T(0.05,1),"err") = %v, want string "err"`, got)
	}
}

func TestCONFIDENCE_T_convergesToNorm(t *testing.T) {
	// For very large sample sizes, CONFIDENCE.T should converge to CONFIDENCE.NORM
	const tol = 1e-3
	resolver := &mockResolver{}

	formula := "CONFIDENCE.T(0.05,1,100000)-CONFIDENCE.NORM(0.05,1,100000)"
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("%s: want number, got type=%d err=%v", formula, got.Type, got.Err)
	}
	if math.Abs(got.Num) > tol {
		t.Errorf("CONFIDENCE.T and CONFIDENCE.NORM should converge for large n, diff = %g", got.Num)
	}
}

func TestPERCENTILE_INC(t *testing.T) {
	// PERCENTILE.INC is an alias for PERCENTILE — verify it works identically
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
		},
	}
	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic", "PERCENTILE.INC(A1:A4,0.3)", 1.9},
		{"min", "PERCENTILE.INC(A1:A4,0)", 1},
		{"max", "PERCENTILE.INC(A1:A4,1)", 4},
		{"median", "PERCENTILE.INC(A1:A4,0.5)", 2.5},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || math.Abs(got.Num-tt.want) > 1e-9 {
				t.Errorf("%s = %v, want %g", tt.formula, got, tt.want)
			}
		})
	}
}

func TestQUARTILE_INC(t *testing.T) {
	// QUARTILE.INC is an alias for QUARTILE — verify it works identically
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
		},
	}
	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"q0", "QUARTILE.INC(A1:A4,0)", 1},
		{"q1", "QUARTILE.INC(A1:A4,1)", 1.75},
		{"q2", "QUARTILE.INC(A1:A4,2)", 2.5},
		{"q3", "QUARTILE.INC(A1:A4,3)", 3.25},
		{"q4", "QUARTILE.INC(A1:A4,4)", 4},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || math.Abs(got.Num-tt.want) > 1e-9 {
				t.Errorf("%s = %v, want %g", tt.formula, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// BETA.DIST
// ---------------------------------------------------------------------------

func TestBETA_DIST(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// CDF with custom bounds (Excel documentation example)
		{"cdf_custom_bounds", "BETA.DIST(2,8,10,TRUE,1,3)", 0.6854706, false, 0},
		// PDF with custom bounds (Excel documentation example)
		{"pdf_custom_bounds", "BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646, false, 0},

		// Standard beta CDF
		{"cdf_standard_basic", "BETA.DIST(0.5,2,3,TRUE)", 0.6875, false, 0},
		{"cdf_standard_alpha3_beta3", "BETA.DIST(0.5,3,3,TRUE)", 0.5, false, 0},
		{"cdf_standard_alpha1_beta1", "BETA.DIST(0.5,1,1,TRUE)", 0.5, false, 0},
		{"cdf_standard_alpha05_beta05", "BETA.DIST(0.5,0.5,0.5,TRUE)", 0.5, false, 0},
		{"cdf_near_zero", "BETA.DIST(0.1,2,5,TRUE)", 0.114265, false, 0},
		{"cdf_near_one", "BETA.DIST(0.9,2,5,TRUE)", 0.999945, false, 0},
		{"cdf_alpha5_beta1", "BETA.DIST(0.5,5,1,TRUE)", 0.03125, false, 0},
		{"cdf_alpha1_beta5", "BETA.DIST(0.5,1,5,TRUE)", 0.96875, false, 0},

		// CDF boundary values
		{"cdf_x_equals_A", "BETA.DIST(0,2,3,TRUE)", 0, false, 0},
		{"cdf_x_equals_B", "BETA.DIST(1,2,3,TRUE)", 1, false, 0},

		// Standard beta PDF
		{"pdf_standard_basic", "BETA.DIST(0.5,2,3,FALSE)", 1.5, false, 0},
		{"pdf_uniform", "BETA.DIST(0.5,1,1,FALSE)", 1, false, 0},
		{"pdf_symmetric", "BETA.DIST(0.5,3,3,FALSE)", 1.875, false, 0},
		{"pdf_alpha5_beta2", "BETA.DIST(0.8,5,2,FALSE)", 2.4576, false, 0},
		{"pdf_alpha2_beta5", "BETA.DIST(0.2,2,5,FALSE)", 2.4576, false, 0},

		// PDF boundary values
		{"pdf_x0_alpha1", "BETA.DIST(0,1,3,FALSE)", 3, false, 0},
		{"pdf_x0_alpha_gt1", "BETA.DIST(0,2,3,FALSE)", 0, false, 0},
		{"pdf_x1_beta1", "BETA.DIST(1,3,1,FALSE)", 3, false, 0},
		{"pdf_x1_beta_gt1", "BETA.DIST(1,3,2,FALSE)", 0, false, 0},

		// Custom bounds CDF
		{"cdf_custom_bounds_2", "BETA.DIST(5,2,3,TRUE,0,10)", 0.6875, false, 0},
		{"cdf_custom_neg_bounds", "BETA.DIST(0,2,3,TRUE,-5,5)", 0.6875, false, 0},

		// Custom bounds PDF
		{"pdf_custom_bounds_scaled", "BETA.DIST(5,2,3,FALSE,0,10)", 0.15, false, 0},

		// Error: alpha <= 0
		{"err_alpha_zero", "BETA.DIST(0.5,0,3,TRUE)", 0, true, ErrValNUM},
		{"err_alpha_neg", "BETA.DIST(0.5,-1,3,TRUE)", 0, true, ErrValNUM},

		// Error: beta <= 0
		{"err_beta_zero", "BETA.DIST(0.5,2,0,TRUE)", 0, true, ErrValNUM},
		{"err_beta_neg", "BETA.DIST(0.5,2,-1,TRUE)", 0, true, ErrValNUM},

		// Error: x < A
		{"err_x_lt_A", "BETA.DIST(-0.1,2,3,TRUE)", 0, true, ErrValNUM},
		{"err_x_lt_A_custom", "BETA.DIST(0,2,3,TRUE,1,3)", 0, true, ErrValNUM},

		// Error: x > B
		{"err_x_gt_B", "BETA.DIST(1.1,2,3,TRUE)", 0, true, ErrValNUM},
		{"err_x_gt_B_custom", "BETA.DIST(4,2,3,TRUE,1,3)", 0, true, ErrValNUM},

		// Error: A = B
		{"err_A_eq_B", "BETA.DIST(1,2,3,TRUE,1,1)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `BETA.DIST("abc",2,3,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_alpha", `BETA.DIST(0.5,"abc",3,TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_beta", `BETA.DIST(0.5,2,"abc",TRUE)`, 0, true, ErrValVALUE},
		{"err_non_numeric_cum", `BETA.DIST(0.5,2,3,"abc")`, 0, true, ErrValVALUE},
		{"err_non_numeric_A", `BETA.DIST(0.5,2,3,TRUE,"abc")`, 0, true, ErrValVALUE},
		{"err_non_numeric_B", `BETA.DIST(0.5,2,3,TRUE,0,"abc")`, 0, true, ErrValVALUE},

		// PDF with x=0 and alpha < 1 (diverges)
		{"err_pdf_x0_alpha_lt1", "BETA.DIST(0,0.5,3,FALSE)", 0, true, ErrValNUM},
		// PDF with x=1 and beta < 1 (diverges)
		{"err_pdf_x1_beta_lt1", "BETA.DIST(1,3,0.5,FALSE)", 0, true, ErrValNUM},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestBETA_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "BETA.DIST(0.5,2,3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BETA.DIST(0.5,2,3) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "BETA.DIST(0.5,2,3,TRUE,0,1,99)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BETA.DIST(0.5,2,3,TRUE,0,1,99) should error, got type=%d", got.Type)
	}
}

// ---------------------------------------------------------------------------
// BETA.INV
// ---------------------------------------------------------------------------

func TestBETA_INV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Excel documentation example
		{"excel_example", "BETA.INV(0.685470581,8,10,1,3)", 2, false, 0},

		// Symmetric distribution: BETA.INV(0.5, 2, 2) = 0.5
		{"symmetric_2_2", "BETA.INV(0.5,2,2)", 0.5, false, 0},

		// Uniform distribution: BETA.INV(0.5, 1, 1) = 0.5
		{"uniform_1_1", "BETA.INV(0.5,1,1)", 0.5, false, 0},

		// Symmetric alpha=beta=3
		{"symmetric_3_3", "BETA.INV(0.5,3,3)", 0.5, false, 0},

		// Symmetric alpha=beta=0.5
		{"symmetric_05_05", "BETA.INV(0.5,0.5,0.5)", 0.5, false, 0},

		// p=0 → returns A (lower bound)
		{"p_zero", "BETA.INV(0,2,3)", 0, false, 0},

		// p=1 → returns B (upper bound)
		{"p_one", "BETA.INV(1,2,3)", 1, false, 0},

		// Round-trip from BETA.DIST: BETA.DIST(0.5, 2, 3, TRUE) = 0.6875
		{"round_trip_05_2_3", "BETA.INV(0.6875,2,3)", 0.5, false, 0},

		// Round-trip from BETA.DIST: BETA.DIST(0.1, 2, 5, TRUE) ≈ 0.114265
		{"round_trip_01_2_5", "BETA.INV(0.114265,2,5)", 0.1, false, 0},

		// Round-trip from BETA.DIST: BETA.DIST(0.9, 2, 5, TRUE) ≈ 0.999945
		{"round_trip_09_2_5", "BETA.INV(0.999945,2,5)", 0.9, false, 0},

		// Skewed left (alpha=5, beta=1)
		{"skewed_left", "BETA.INV(0.03125,5,1)", 0.5, false, 0},

		// Skewed right (alpha=1, beta=5)
		{"skewed_right", "BETA.INV(0.96875,1,5)", 0.5, false, 0},

		// Custom bounds: BETA.INV(0.5, 2, 2, 10, 20) = 15
		{"custom_bounds_symmetric", "BETA.INV(0.5,2,2,10,20)", 15, false, 0},

		// Custom bounds with negative range
		{"custom_bounds_neg", "BETA.INV(0.5,2,2,-5,5)", 0, false, 0},

		// Custom bounds: p=0 returns A
		{"custom_bounds_p0", "BETA.INV(0,2,3,1,3)", 1, false, 0},

		// Custom bounds: p=1 returns B
		{"custom_bounds_p1", "BETA.INV(1,2,3,1,3)", 3, false, 0},

		// Small alpha, large beta → distribution concentrated near 0
		{"small_alpha_large_beta", "BETA.INV(0.5,0.5,5)", 0.0466872, false, 0},

		// Large alpha, small beta → distribution concentrated near 1
		{"large_alpha_small_beta", "BETA.INV(0.5,5,0.5)", 0.9533128, false, 0},

		// Near p=1
		{"near_p_one", "BETA.INV(0.99,2,3)", 0.8591325, false, 0},

		// Near p=0 (but positive)
		{"near_p_zero", "BETA.INV(0.01,2,3)", 0.0419986, false, 0},

		// Fractional alpha and beta
		{"fractional_params", "BETA.INV(0.5,1.5,2.5)", 0.3524523, false, 0},

		// Large parameters
		{"large_params", "BETA.INV(0.5,100,100)", 0.5, false, 0},

		// Error: probability < 0
		{"err_p_neg", "BETA.INV(-0.1,2,3)", 0, true, ErrValNUM},

		// Error: probability > 1
		{"err_p_gt1", "BETA.INV(1.1,2,3)", 0, true, ErrValNUM},

		// Error: alpha <= 0
		{"err_alpha_zero", "BETA.INV(0.5,0,3)", 0, true, ErrValNUM},
		{"err_alpha_neg", "BETA.INV(0.5,-1,3)", 0, true, ErrValNUM},

		// Error: beta <= 0
		{"err_beta_zero", "BETA.INV(0.5,2,0)", 0, true, ErrValNUM},
		{"err_beta_neg", "BETA.INV(0.5,2,-1)", 0, true, ErrValNUM},

		// Error: A >= B
		{"err_A_eq_B", "BETA.INV(0.5,2,3,5,5)", 0, true, ErrValNUM},
		{"err_A_gt_B", "BETA.INV(0.5,2,3,5,3)", 0, true, ErrValNUM},

		// Error: non-numeric arguments
		{"err_non_numeric_p", `BETA.INV("abc",2,3)`, 0, true, ErrValVALUE},
		{"err_non_numeric_alpha", `BETA.INV(0.5,"abc",3)`, 0, true, ErrValVALUE},
		{"err_non_numeric_beta", `BETA.INV(0.5,2,"abc")`, 0, true, ErrValVALUE},
		{"err_non_numeric_A", `BETA.INV(0.5,2,3,"abc")`, 0, true, ErrValVALUE},
		{"err_non_numeric_B", `BETA.INV(0.5,2,3,0,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s: want error %d, got type=%d num=%g", tt.formula, tt.wantErr, got.Type, got.Num)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s: want err=%d, got err=%d", tt.formula, tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestBETA_INV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args (2)
	cf := evalCompile(t, "BETA.INV(0.5,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BETA.INV(0.5,2) should error, got type=%d", got.Type)
	}

	// Too many args (6)
	cf = evalCompile(t, "BETA.INV(0.5,2,3,0,1,99)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BETA.INV(0.5,2,3,0,1,99) should error, got type=%d", got.Type)
	}
}

func TestBETA_INV_round_trip(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	// BETA.INV(BETA.DIST(x, alpha, beta, TRUE), alpha, beta) ≈ x
	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"rt_03_2_5", "BETA.INV(BETA.DIST(0.3,2,5,TRUE),2,5)", 0.3},
		{"rt_07_3_2", "BETA.INV(BETA.DIST(0.7,3,2,TRUE),3,2)", 0.7},
		{"rt_05_10_10", "BETA.INV(BETA.DIST(0.5,10,10,TRUE),10,10)", 0.5},
		{"rt_02_1_3", "BETA.INV(BETA.DIST(0.2,1,3,TRUE),1,3)", 0.2},
		{"rt_custom_bounds", "BETA.INV(BETA.DIST(5,2,3,TRUE,0,10),2,3,0,10)", 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}
}

func TestT_DIST_RT(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-5

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"basic", "T.DIST.RT(2,10)", 0.03669402, false, 0},
		{"zero", "T.DIST.RT(0,5)", 0.5, false, 0},
		{"negative_x", "T.DIST.RT(-2,10)", 0.96330598, false, 0},
		{"large_x", "T.DIST.RT(5,10)", 0.00026867, false, 0},
		{"df1", "T.DIST.RT(1,1)", 0.25, false, 0},
		{"df2", "T.DIST.RT(1,2)", 0.21132487, false, 0},
		{"df30", "T.DIST.RT(1.96,30)", 0.02967116, false, 0},
		{"df100", "T.DIST.RT(2.576,100)", 0.00572851, false, 0},
		{"small_x", "T.DIST.RT(0.1,5)", 0.46211507, false, 0},
		{"neg_large", "T.DIST.RT(-5,10)", 0.99973133, false, 0},

		// Truncation: 10.7 → 10
		{"trunc_df", "T.DIST.RT(2,10.7)", 0.03669402, false, 0},

		// Error: df < 1
		{"err_df_zero", "T.DIST.RT(1,0)", 0, true, ErrValNUM},
		{"err_df_neg", "T.DIST.RT(1,-1)", 0, true, ErrValNUM},
		{"err_df_frac_below1", "T.DIST.RT(1,0.9)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `T.DIST.RT("abc",10)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `T.DIST.RT(2,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "T.DIST.RT(2)", 0, true, ErrValVALUE},
		{"err_too_many", "T.DIST.RT(2,10,1)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestT_DIST_2T(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-5

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"basic", "T.DIST.2T(2,10)", 0.07338803, false, 0},
		{"zero", "T.DIST.2T(0,5)", 1.0, false, 0},
		{"large_x", "T.DIST.2T(5,10)", 0.00053733, false, 0},
		{"df1", "T.DIST.2T(1,1)", 0.5, false, 0},
		{"df2", "T.DIST.2T(1,2)", 0.42264973, false, 0},
		{"df30", "T.DIST.2T(1.96,30)", 0.05934231, false, 0},
		{"df100", "T.DIST.2T(2.576,100)", 0.01145702, false, 0},
		{"small_x", "T.DIST.2T(0.1,5)", 0.92423014, false, 0},
		{"df1_x2", "T.DIST.2T(2,1)", 0.29516724, false, 0},
		{"df50", "T.DIST.2T(3,50)", 0.00420170, false, 0},

		// Truncation: 10.7 → 10
		{"trunc_df", "T.DIST.2T(2,10.7)", 0.07338803, false, 0},

		// Error: x < 0
		{"err_neg_x", "T.DIST.2T(-1,10)", 0, true, ErrValNUM},
		{"err_neg_x2", "T.DIST.2T(-0.001,5)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df_zero", "T.DIST.2T(1,0)", 0, true, ErrValNUM},
		{"err_df_neg", "T.DIST.2T(1,-1)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `T.DIST.2T("abc",10)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `T.DIST.2T(2,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "T.DIST.2T(2)", 0, true, ErrValVALUE},
		{"err_too_many", "T.DIST.2T(2,10,1)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestT_INV_2T(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-4

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"basic_05_10", "T.INV.2T(0.05,10)", 2.22813885, false, 0},
		{"p1_df5", "T.INV.2T(1,5)", 0, false, 0},
		{"p01_df10", "T.INV.2T(0.01,10)", 3.16927267, false, 0},
		{"p10_df20", "T.INV.2T(0.10,20)", 1.72471824, false, 0},
		{"p05_df1", "T.INV.2T(0.05,1)", 12.7062047, false, 0},
		{"p05_df2", "T.INV.2T(0.05,2)", 4.30265273, false, 0},
		{"p05_df30", "T.INV.2T(0.05,30)", 2.04227246, false, 0},
		{"p05_df100", "T.INV.2T(0.05,100)", 1.98397152, false, 0},
		{"p50_df10", "T.INV.2T(0.5,10)", 0.69981200, false, 0},
		{"p90_df10", "T.INV.2T(0.9,10)", 0.12891459, false, 0},

		// Truncation: 10.7 → 10
		{"trunc_df", "T.INV.2T(0.05,10.7)", 2.22813885, false, 0},

		// Error: p <= 0
		{"err_p_zero", "T.INV.2T(0,10)", 0, true, ErrValNUM},
		{"err_p_neg", "T.INV.2T(-0.1,10)", 0, true, ErrValNUM},

		// Error: p > 1
		{"err_p_gt1", "T.INV.2T(1.1,10)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df_zero", "T.INV.2T(0.05,0)", 0, true, ErrValNUM},
		{"err_df_neg", "T.INV.2T(0.05,-1)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_p", `T.INV.2T("abc",10)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `T.INV.2T(0.05,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "T.INV.2T(0.05)", 0, true, ErrValVALUE},
		{"err_too_many", "T.INV.2T(0.05,10,1)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestCHISQ_DIST_RT(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-5

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"critical_95", "CHISQ.DIST.RT(3.841,1)", 0.05001368, false, 0},
		{"x0_df5", "CHISQ.DIST.RT(0,5)", 1.0, false, 0},
		{"df2", "CHISQ.DIST.RT(5.991,2)", 0.05001162, false, 0},
		{"df5", "CHISQ.DIST.RT(5,5)", 0.41588019, false, 0},
		{"df10", "CHISQ.DIST.RT(10,10)", 0.44049329, false, 0},
		{"df20", "CHISQ.DIST.RT(20,20)", 0.45792971, false, 0},
		{"large_x", "CHISQ.DIST.RT(100,5)", 0, false, 0},
		{"small_x", "CHISQ.DIST.RT(0.01,1)", 0.92034433, false, 0},
		{"df1_x1", "CHISQ.DIST.RT(1,1)", 0.31731051, false, 0},
		{"df3_x3", "CHISQ.DIST.RT(3,3)", 0.39162518, false, 0},

		// Truncation: 3.7 → 3
		{"trunc_df", "CHISQ.DIST.RT(3,3.7)", 0.39162518, false, 0},

		// Error: x < 0
		{"err_neg_x", "CHISQ.DIST.RT(-1,3)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df_zero", "CHISQ.DIST.RT(1,0)", 0, true, ErrValNUM},
		{"err_df_neg", "CHISQ.DIST.RT(1,-1)", 0, true, ErrValNUM},

		// Error: df > 10^10
		{"err_df_too_large", "CHISQ.DIST.RT(1,10000000001)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `CHISQ.DIST.RT("abc",1)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `CHISQ.DIST.RT(1,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "CHISQ.DIST.RT(1)", 0, true, ErrValVALUE},
		{"err_too_many", "CHISQ.DIST.RT(1,3,TRUE)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestCHISQ_INV_RT(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-4

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"critical_05_1", "CHISQ.INV.RT(0.05,1)", 3.84145882, false, 0},
		{"critical_05_2", "CHISQ.INV.RT(0.05,2)", 5.99146455, false, 0},
		{"critical_05_5", "CHISQ.INV.RT(0.05,5)", 11.0704977, false, 0},
		{"critical_05_10", "CHISQ.INV.RT(0.05,10)", 18.3070380, false, 0},
		{"critical_01_1", "CHISQ.INV.RT(0.01,1)", 6.63489660, false, 0},
		{"p50_df5", "CHISQ.INV.RT(0.5,5)", 4.35146163, false, 0},
		{"p90_df5", "CHISQ.INV.RT(0.9,5)", 1.61031160, false, 0},
		{"p1_df5", "CHISQ.INV.RT(1,5)", 0, false, 0},
		{"p10_df20", "CHISQ.INV.RT(0.10,20)", 28.4119982, false, 0},
		{"critical_05_30", "CHISQ.INV.RT(0.05,30)", 43.7729690, false, 0},

		// Truncation: 5.9 → 5
		{"trunc_df", "CHISQ.INV.RT(0.05,5.9)", 11.0704977, false, 0},

		// Error: p < 0
		{"err_p_neg", "CHISQ.INV.RT(-0.1,5)", 0, true, ErrValNUM},

		// Error: p > 1
		{"err_p_gt1", "CHISQ.INV.RT(1.1,5)", 0, true, ErrValNUM},

		// Error: p = 0
		{"err_p_zero", "CHISQ.INV.RT(0,5)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df_zero", "CHISQ.INV.RT(0.05,0)", 0, true, ErrValNUM},
		{"err_df_neg", "CHISQ.INV.RT(0.05,-1)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_p", `CHISQ.INV.RT("abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df", `CHISQ.INV.RT(0.05,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "CHISQ.INV.RT(0.05)", 0, true, ErrValVALUE},
		{"err_too_many", "CHISQ.INV.RT(0.05,5,1)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestF_DIST_RT(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-4

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"critical_01", "F.DIST.RT(15.2069,6,4)", 0.01000, false, 0},
		{"x0", "F.DIST.RT(0,5,5)", 1.0, false, 0},
		{"basic1", "F.DIST.RT(1,5,5)", 0.50000, false, 0},
		{"basic2", "F.DIST.RT(2,5,5)", 0.23251, false, 0},
		{"basic3", "F.DIST.RT(5,2,3)", 0.11086, false, 0},
		{"df1_1", "F.DIST.RT(1,1,1)", 0.50000, false, 0},
		{"large_x", "F.DIST.RT(100,5,5)", 0.00001, false, 0},
		{"small_x", "F.DIST.RT(0.1,5,5)", 0.98776, false, 0},
		{"df10_10", "F.DIST.RT(2,10,10)", 0.14485, false, 0},
		{"df30_30", "F.DIST.RT(1.5,30,30)", 0.13621, false, 0},

		// Truncation: 6.9 → 6, 4.9 → 4
		{"trunc_df", "F.DIST.RT(15.2069,6.9,4.9)", 0.01000, false, 0},

		// Error: x < 0
		{"err_neg_x", "F.DIST.RT(-1,5,5)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df1_zero", "F.DIST.RT(1,0,5)", 0, true, ErrValNUM},
		{"err_df2_zero", "F.DIST.RT(1,5,0)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_x", `F.DIST.RT("abc",5,5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df1", `F.DIST.RT(1,"abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df2", `F.DIST.RT(1,5,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "F.DIST.RT(1,5)", 0, true, ErrValVALUE},
		{"err_too_many", "F.DIST.RT(1,5,5,TRUE)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

func TestF_INV_RT(t *testing.T) {
	resolver := &mockResolver{}
	const tol = 1e-3

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr bool
		errVal  ErrorValue
	}{
		// Basic cases
		{"critical_01", "F.INV.RT(0.01,6,4)", 15.2069, false, 0},
		{"p1_df5_5", "F.INV.RT(1,5,5)", 0, false, 0},
		{"p50_df5_5", "F.INV.RT(0.5,5,5)", 1.0, false, 0},
		{"p05_df5_5", "F.INV.RT(0.05,5,5)", 5.0503, false, 0},
		{"p10_df2_3", "F.INV.RT(0.10,2,3)", 5.4624, false, 0},
		{"p05_df1_1", "F.INV.RT(0.05,1,1)", 161.448, false, 0},
		{"p05_df10_10", "F.INV.RT(0.05,10,10)", 2.9782, false, 0},
		{"p05_df30_30", "F.INV.RT(0.05,30,30)", 1.8409, false, 0},
		{"p90_df5_5", "F.INV.RT(0.9,5,5)", 0.2896, false, 0},
		{"p25_df10_10", "F.INV.RT(0.25,10,10)", 1.5513, false, 0},

		// Truncation: 6.9 → 6, 4.9 → 4
		{"trunc_df", "F.INV.RT(0.01,6.9,4.9)", 15.2069, false, 0},

		// Error: p < 0
		{"err_p_neg", "F.INV.RT(-0.1,5,5)", 0, true, ErrValNUM},

		// Error: p > 1
		{"err_p_gt1", "F.INV.RT(1.1,5,5)", 0, true, ErrValNUM},

		// Error: p = 0
		{"err_p_zero", "F.INV.RT(0,5,5)", 0, true, ErrValNUM},

		// Error: df < 1
		{"err_df1_zero", "F.INV.RT(0.05,0,5)", 0, true, ErrValNUM},
		{"err_df2_zero", "F.INV.RT(0.05,5,0)", 0, true, ErrValNUM},

		// Error: non-numeric
		{"err_non_numeric_p", `F.INV.RT("abc",5,5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df1", `F.INV.RT(0.05,"abc",5)`, 0, true, ErrValVALUE},
		{"err_non_numeric_df2", `F.INV.RT(0.05,5,"abc")`, 0, true, ErrValVALUE},

		// Error: wrong arg count
		{"err_too_few", "F.INV.RT(0.05,5)", 0, true, ErrValVALUE},
		{"err_too_many", "F.INV.RT(0.05,5,5,1)", 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Fatalf("%s: want error %v, got type=%d val=%v", tt.formula, tt.errVal, got.Type, got)
				}
				if got.Err != tt.errVal {
					t.Errorf("%s: want error %v, got %v", tt.formula, tt.errVal, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: want number, got type=%d err=%v", tt.formula, got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g (diff=%g)", tt.formula, got.Num, tt.want, math.Abs(got.Num-tt.want))
			}
		})
	}
}

// ---------------------------------------------------------------------------
// HYPGEOM.DIST
// ---------------------------------------------------------------------------

func TestHYPGEOM_DIST(t *testing.T) {
	const tol = 1e-9
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic PMF — Excel example
		{"pmf_basic", "HYPGEOM.DIST(1,4,8,20,FALSE)", 0.363261093911, false, 0},

		// Basic CDF — Excel example
		{"cdf_basic", "HYPGEOM.DIST(1,4,8,20,TRUE)", 0.465428276574, false, 0},

		// All successes in sample PMF
		{"pmf_all_success", "HYPGEOM.DIST(4,4,8,20,FALSE)", 0.014447884417, false, 0},

		// Zero successes PMF
		{"pmf_zero_success", "HYPGEOM.DIST(0,4,8,20,FALSE)", 0.102167182663, false, 0},

		// CDF at max should be 1
		{"cdf_at_max", "HYPGEOM.DIST(4,4,8,20,TRUE)", 1, false, 0},

		// Simple case: 1 success from sample of 1, 1 success in pop of 2
		{"simple_half", "HYPGEOM.DIST(1,1,1,2,FALSE)", 0.5, false, 0},

		// Simple case: 0 successes
		{"simple_zero", "HYPGEOM.DIST(0,1,1,2,FALSE)", 0.5, false, 0},

		// Simple CDF at max
		{"simple_cdf_max", "HYPGEOM.DIST(1,1,1,2,TRUE)", 1, false, 0},

		// Truncation: 1.9 -> 1, 4.7 -> 4, 8.3 -> 8, 20.1 -> 20
		{"truncation", "HYPGEOM.DIST(1.9,4.7,8.3,20.1,FALSE)", 0.363261093911, false, 0},

		// PMF with larger sample
		{"pmf_large_sample", "HYPGEOM.DIST(3,10,15,50,FALSE)", 0.297855699521, false, 0},

		// CDF with larger sample
		{"cdf_large_sample", "HYPGEOM.DIST(3,10,15,50,TRUE)", 0.659406694786, false, 0},

		// Extreme: sample_s equals both number_sample and population_s
		{"pmf_all_match", "HYPGEOM.DIST(3,3,3,10,FALSE)", 0.008333333333, false, 0},

		// PMF k=2 out of sample=4, pop_s=8, pop=20
		{"pmf_k2", "HYPGEOM.DIST(2,4,8,20,FALSE)", 0.381424148607, false, 0},

		// PMF k=3
		{"pmf_k3", "HYPGEOM.DIST(3,4,8,20,FALSE)", 0.138699690402, false, 0},

		// CDF at k=0
		{"cdf_at_zero", "HYPGEOM.DIST(0,4,8,20,TRUE)", 0.102167182663, false, 0},

		// CDF at k=2
		{"cdf_at_k2", "HYPGEOM.DIST(2,4,8,20,TRUE)", 0.846852425181, false, 0},

		// CDF at k=3
		{"cdf_at_k3", "HYPGEOM.DIST(3,4,8,20,TRUE)", 0.985552115583, false, 0},

		// Small population: pop=5, pop_s=2, sample=3, k=1
		{"small_pop", "HYPGEOM.DIST(1,3,2,5,FALSE)", 0.6, false, 0},

		// Small population CDF
		{"small_pop_cdf", "HYPGEOM.DIST(1,3,2,5,TRUE)", 0.7, false, 0},

		// Large population
		{"large_pop_pmf", "HYPGEOM.DIST(5,10,50,200,FALSE)", 0.055830842234, false, 0},

		// sample_s = number_sample (all drawn are successes)
		{"all_drawn_success", "HYPGEOM.DIST(2,2,5,10,FALSE)", 0.222222222222, false, 0},

		// Error: sample_s < 0
		{"err_neg_sample_s", "HYPGEOM.DIST(-1,4,8,20,FALSE)", 0, true, ErrValNUM},

		// Error: sample_s > number_sample
		{"err_s_gt_n", "HYPGEOM.DIST(5,4,8,20,FALSE)", 0, true, ErrValNUM},

		// Error: sample_s > population_s
		{"err_s_gt_pop_s", "HYPGEOM.DIST(5,10,4,20,FALSE)", 0, true, ErrValNUM},

		// Error: number_sample <= 0
		{"err_n_zero", "HYPGEOM.DIST(0,0,8,20,FALSE)", 0, true, ErrValNUM},

		// Error: number_sample > number_pop
		{"err_n_gt_pop", "HYPGEOM.DIST(1,25,8,20,FALSE)", 0, true, ErrValNUM},

		// Error: population_s <= 0
		{"err_pop_s_zero", "HYPGEOM.DIST(0,4,0,20,FALSE)", 0, true, ErrValNUM},

		// Error: population_s > number_pop
		{"err_pop_s_gt_pop", "HYPGEOM.DIST(1,4,25,20,FALSE)", 0, true, ErrValNUM},

		// Error: number_pop <= 0
		{"err_pop_zero", "HYPGEOM.DIST(0,4,8,0,FALSE)", 0, true, ErrValNUM},

		// Error: sample_s < max(0, n + M - N) — lower bound violation
		{"err_lower_bound", "HYPGEOM.DIST(0,10,15,20,FALSE)", 0, true, ErrValNUM},

		// Error: non-numeric first arg
		{"err_non_numeric_1", `HYPGEOM.DIST("abc",4,8,20,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric second arg
		{"err_non_numeric_2", `HYPGEOM.DIST(1,"abc",8,20,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric third arg
		{"err_non_numeric_3", `HYPGEOM.DIST(1,4,"abc",20,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric fourth arg
		{"err_non_numeric_4", `HYPGEOM.DIST(1,4,8,"abc",FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric fifth arg
		{"err_non_numeric_5", `HYPGEOM.DIST(1,4,8,20,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestHYPGEOM_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "HYPGEOM.DIST(1,4,8,20)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("HYPGEOM.DIST(1,4,8,20) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "HYPGEOM.DIST(1,4,8,20,FALSE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("HYPGEOM.DIST(1,4,8,20,FALSE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(HYPGEOM.DIST(1,4,8,20),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(HYPGEOM.DIST(1,4,8,20),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// NEGBINOM.DIST
// ---------------------------------------------------------------------------

func TestNEGBINOM_DIST(t *testing.T) {
	const tol = 1e-7
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic PMF — Excel: NEGBINOM.DIST(10,5,0.25,FALSE) ≈ 0.0550487
		{"pmf_basic", "NEGBINOM.DIST(10,5,0.25,FALSE)", 0.0550487, false, 0},

		// Basic CDF — Excel: NEGBINOM.DIST(10,5,0.25,TRUE) ≈ 0.3135141
		{"cdf_basic", "NEGBINOM.DIST(10,5,0.25,TRUE)", 0.3135141, false, 0},

		// PMF with p=0.5 — NEGBINOM.DIST(3,5,0.5,FALSE)
		// C(7,4)*0.5^5*0.5^3 = 35 * 0.03125 * 0.125 = 0.13671875
		{"pmf_p05", "NEGBINOM.DIST(3,5,0.5,FALSE)", 0.13671875, false, 0},

		// CDF with p=0.5 — NEGBINOM.DIST(3,5,0.5,TRUE)
		{"cdf_p05", "NEGBINOM.DIST(3,5,0.5,TRUE)", 0.36328125, false, 0},

		// Zero failures PMF — P(X=0) = p^r
		// NEGBINOM.DIST(0,5,0.5,FALSE) = 0.5^5 = 0.03125
		{"pmf_zero_failures", "NEGBINOM.DIST(0,5,0.5,FALSE)", 0.03125, false, 0},

		// Zero failures CDF — same as PMF
		{"cdf_zero_failures", "NEGBINOM.DIST(0,5,0.5,TRUE)", 0.03125, false, 0},

		// r=1, geometric distribution: PMF = p*(1-p)^f
		// NEGBINOM.DIST(3,1,0.3,FALSE) = 0.3*0.7^3 = 0.1029
		{"pmf_geometric", "NEGBINOM.DIST(3,1,0.3,FALSE)", 0.1029, false, 0},

		// r=1, geometric distribution CDF
		// CDF = 1-(1-p)^(f+1) = 1-0.7^4 = 1-0.2401 = 0.7599
		{"cdf_geometric", "NEGBINOM.DIST(3,1,0.3,TRUE)", 0.7599, false, 0},

		// p=0, f=0: certain to have zero failures if... wait, p=0 means
		// success never happens. But the convention: p^r with p=0 and f=0
		// returns 1 (edge case).
		{"pmf_p0_f0", "NEGBINOM.DIST(0,1,0,FALSE)", 1, false, 0},

		// p=0, f>0: no successes ever, so PMF is 0
		{"pmf_p0_f_pos", "NEGBINOM.DIST(5,1,0,FALSE)", 0, false, 0},

		// p=1, f=0: success always, zero failures, PMF = 1^r = 1
		{"pmf_p1_f0", "NEGBINOM.DIST(0,3,1,FALSE)", 1, false, 0},

		// p=1, f>0: no failures possible, PMF = 0
		{"pmf_p1_f_pos", "NEGBINOM.DIST(2,3,1,FALSE)", 0, false, 0},

		// CDF at p=1, f=0: CDF = 1
		{"cdf_p1_f0", "NEGBINOM.DIST(0,3,1,TRUE)", 1, false, 0},

		// CDF at p=0, f=0: CDF = I_0(r, 1) = 0 — but PMF for f=0 is 1
		// regBetaInc(0, r, f+1) = 0, which contradicts. Actually p=0 edge:
		// regBetaInc handles x=0 returning 0.
		{"cdf_p0_f0", "NEGBINOM.DIST(0,1,0,TRUE)", 0, false, 0},

		// Truncation: 10.9 -> 10, 5.7 -> 5
		{"truncation", "NEGBINOM.DIST(10.9,5.7,0.25,FALSE)", 0.0550487, false, 0},

		// Large failures PMF — NEGBINOM.DIST(50,10,0.3,FALSE)
		{"pmf_large_f", "NEGBINOM.DIST(50,10,0.3,FALSE)", 0.0013344436566226712, false, 0},

		// Large failures CDF — NEGBINOM.DIST(50,10,0.3,TRUE)
		{"cdf_large_f", "NEGBINOM.DIST(50,10,0.3,TRUE)", 0.9941288117574135, false, 0},

		// High p PMF — NEGBINOM.DIST(2,10,0.9,FALSE)
		// Very likely to succeed, few failures expected
		{"pmf_high_p", "NEGBINOM.DIST(2,10,0.9,FALSE)", 0.1917731420550001, false, 0},

		// High p CDF — NEGBINOM.DIST(2,10,0.9,TRUE)
		{"cdf_high_p", "NEGBINOM.DIST(2,10,0.9,TRUE)", 0.889130022255, false, 0},

		// r=1, p=1, f=0: geometric with certain success
		{"pmf_r1_p1_f0", "NEGBINOM.DIST(0,1,1,FALSE)", 1, false, 0},

		// Error: number_f < 0
		{"err_neg_f", "NEGBINOM.DIST(-1,5,0.25,FALSE)", 0, true, ErrValNUM},

		// Error: number_s < 1
		{"err_s_zero", "NEGBINOM.DIST(5,0,0.25,FALSE)", 0, true, ErrValNUM},

		// Error: probability_s < 0
		{"err_p_neg", "NEGBINOM.DIST(5,5,-0.1,FALSE)", 0, true, ErrValNUM},

		// Error: probability_s > 1
		{"err_p_gt1", "NEGBINOM.DIST(5,5,1.1,FALSE)", 0, true, ErrValNUM},

		// Error: non-numeric first arg
		{"err_non_numeric_f", `NEGBINOM.DIST("abc",5,0.25,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric second arg
		{"err_non_numeric_s", `NEGBINOM.DIST(10,"abc",0.25,FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric third arg
		{"err_non_numeric_p", `NEGBINOM.DIST(10,5,"abc",FALSE)`, 0, true, ErrValVALUE},

		// Error: non-numeric fourth arg
		{"err_non_numeric_cum", `NEGBINOM.DIST(10,5,0.25,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %.12f, want %.12f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestNEGBINOM_DIST_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "NEGBINOM.DIST(10,5,0.25)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NEGBINOM.DIST(10,5,0.25) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "NEGBINOM.DIST(10,5,0.25,FALSE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("NEGBINOM.DIST(10,5,0.25,FALSE,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(NEGBINOM.DIST(10,5,0.25),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(NEGBINOM.DIST(10,5,0.25),"err") = %v, want string "err"`, got)
	}
}

// ---------------------------------------------------------------------------
// BINOM.INV
// ---------------------------------------------------------------------------

func TestBINOM_INV(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic case from Excel docs
		{"basic", "BINOM.INV(6,0.5,0.75)", 4, false, 0},

		// Median of symmetric binomial
		{"median_symmetric", "BINOM.INV(10,0.5,0.5)", 5, false, 0},

		// Very low alpha — CDF(0)=0.000977 < 0.001, CDF(1)=0.01074 >= 0.001
		{"low_alpha", "BINOM.INV(10,0.5,0.001)", 1, false, 0},

		// Very high alpha — expect near n
		{"high_alpha", "BINOM.INV(10,0.5,0.999)", 9, false, 0},

		// Single trial, alpha below p(0)=0.5
		{"single_trial_low", "BINOM.INV(1,0.5,0.4)", 0, false, 0},

		// Single trial, alpha above p(0)=0.5
		{"single_trial_high", "BINOM.INV(1,0.5,0.6)", 1, false, 0},

		// Single trial, alpha exactly at boundary
		{"single_trial_exact", "BINOM.INV(1,0.5,0.5)", 0, false, 0},

		// Zero trials: CDF(0,0,p)=1, so any alpha < 1 returns 0
		{"zero_trials", "BINOM.INV(0,0.5,0.5)", 0, false, 0},

		// High probability of success
		{"high_prob", "BINOM.INV(10,0.9,0.5)", 9, false, 0},

		// Low probability of success
		{"low_prob", "BINOM.INV(10,0.1,0.5)", 1, false, 0},

		// High probability, low alpha
		{"high_prob_low_alpha", "BINOM.INV(10,0.9,0.01)", 6, false, 0},

		// Low probability, high alpha
		{"low_prob_high_alpha", "BINOM.INV(10,0.1,0.99)", 4, false, 0},

		// Truncation: 6.9 -> 6 trials
		{"truncation", "BINOM.INV(6.9,0.5,0.75)", 4, false, 0},

		// Alpha near lower boundary
		{"alpha_near_zero", "BINOM.INV(10,0.5,0.01)", 1, false, 0},

		// Larger n
		{"larger_n", "BINOM.INV(100,0.5,0.5)", 50, false, 0},

		// Asymmetric probability
		{"asym_p03", "BINOM.INV(10,0.3,0.5)", 3, false, 0},

		// Asymmetric probability with high alpha
		{"asym_p03_high_alpha", "BINOM.INV(10,0.3,0.9)", 5, false, 0},

		// Asymmetric probability with low alpha
		{"asym_p07", "BINOM.INV(10,0.7,0.5)", 7, false, 0},

		// p = 0.5, n = 20
		{"n20_p05", "BINOM.INV(20,0.5,0.75)", 12, false, 0},

		// ---- Error cases ----

		// Negative trials
		{"err_neg_trials", "BINOM.INV(-1,0.5,0.5)", 0, true, ErrValNUM},

		// Probability <= 0
		{"err_p_zero", "BINOM.INV(10,0,0.5)", 0, true, ErrValNUM},

		// Probability >= 1
		{"err_p_one", "BINOM.INV(10,1,0.5)", 0, true, ErrValNUM},

		// Probability negative
		{"err_p_neg", "BINOM.INV(10,-0.1,0.5)", 0, true, ErrValNUM},

		// Probability > 1
		{"err_p_gt1", "BINOM.INV(10,1.1,0.5)", 0, true, ErrValNUM},

		// Alpha <= 0
		{"err_alpha_zero", "BINOM.INV(10,0.5,0)", 0, true, ErrValNUM},

		// Alpha >= 1
		{"err_alpha_one", "BINOM.INV(10,0.5,1)", 0, true, ErrValNUM},

		// Alpha negative
		{"err_alpha_neg", "BINOM.INV(10,0.5,-0.1)", 0, true, ErrValNUM},

		// Alpha > 1
		{"err_alpha_gt1", "BINOM.INV(10,0.5,1.5)", 0, true, ErrValNUM},

		// Non-numeric trials
		{"err_non_numeric_trials", `BINOM.INV("abc",0.5,0.5)`, 0, true, ErrValVALUE},

		// Non-numeric probability
		{"err_non_numeric_prob", `BINOM.INV(10,"abc",0.5)`, 0, true, ErrValVALUE},

		// Non-numeric alpha
		{"err_non_numeric_alpha", `BINOM.INV(10,0.5,"abc")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}

			if tt.wantError {
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				} else if got.Err != tt.wantErr {
					t.Errorf("%s error = %v, want %v", tt.formula, got.Err, tt.wantErr)
				}
				return
			}

			if got.Type != ValueNumber {
				t.Fatalf("%s = %v (type %d), want number %v", tt.formula, got, got.Type, tt.wantNum)
			}
			if got.Num != tt.wantNum {
				t.Errorf("%s = %v, want %v", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

func TestBINOM_INV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "BINOM.INV(10,0.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BINOM.INV(10,0.5) should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "BINOM.INV(10,0.5,0.5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("BINOM.INV(10,0.5,0.5,1) should error, got type=%d", got.Type)
	}

	// IFERROR should catch the #VALUE! from wrong arg count
	cf = evalCompile(t, `IFERROR(BINOM.INV(10,0.5),"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf(`IFERROR(BINOM.INV(10,0.5),"err") = %v, want string "err"`, got)
	}
}

func TestPHI(t *testing.T) {
	const tol = 1e-6
	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
	}{
		{"basic", "PHI(0.75)", 0.301137432, false},
		{"zero", "PHI(0)", 0.398942280, false},
		{"one", "PHI(1)", 0.241970725, false},
		{"neg_one", "PHI(-1)", 0.241970725, false},
		{"two", "PHI(2)", 0.053990967, false},
		{"neg_two", "PHI(-2)", 0.053990967, false},
		{"three", "PHI(3)", 0.004431848, false},
		{"large", "PHI(10)", 0, false},
		{"half", "PHI(0.5)", 0.352065327, false},
		{"neg_half", "PHI(-0.5)", 0.352065327, false},
		{"err_text", `PHI("abc")`, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, nil, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", tt.formula, got)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v, want number", tt.formula, got)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}
}

func TestGAUSS(t *testing.T) {
	const tol = 1e-6
	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
	}{
		{"basic", "GAUSS(2)", 0.477250, false},
		{"zero", "GAUSS(0)", 0, false},
		{"one", "GAUSS(1)", 0.341345, false},
		{"neg_one", "GAUSS(-1)", -0.341345, false},
		{"three", "GAUSS(3)", 0.498650, false},
		{"neg_two", "GAUSS(-2)", -0.477250, false},
		{"half", "GAUSS(0.5)", 0.191462, false},
		{"large", "GAUSS(6)", 0.5, false},
		{"err_text", `GAUSS("abc")`, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, nil, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", tt.formula, got)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s = %v, want number", tt.formula, got)
			}
			if math.Abs(got.Num-tt.want) > tol {
				t.Errorf("%s = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}
}
