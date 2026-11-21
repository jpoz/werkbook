package formula

import "testing"

func TestParseLiterals(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"42", "42"},
		{"1.5", "1.5"},
		{".5", ".5"},
		{"1.5E10", "1.5E10"},
		{"0", "0"},
		{`"hello"`, `"hello"`},
		{`""`, `""`},
		{`"he said ""hi"""`, `"he said \"hi\""`}, // lexer un-doubles quotes; String() uses Go %q
		{"TRUE", "TRUE"},
		{"FALSE", "FALSE"},
		{"true", "TRUE"},
		{"#N/A", "#N/A"},
		{"#DIV/0!", "#DIV/0!"},
		{"#VALUE!", "#VALUE!"},
		{"#REF!", "#REF!"},
		{"#NAME?", "#NAME?"},
		{"#NUM!", "#NUM!"},
		{"#NULL!", "#NULL!"},
		{"#SPILL!", "#SPILL!"},
		{"#CALC!", "#CALC!"},
		{"@A1", "(@ A1)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseArithmetic(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"1+2", "(+ 1 2)"},
		{"2*3", "(* 2 3)"},
		{"10/5", "(/ 10 5)"},
		{"3-1", "(- 3 1)"},
		{"2^3", "(^ 2 3)"},
		{`"a"&"b"`, `(& "a" "b")`},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParsePrecedence(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// * binds tighter than +
		{"2+3*4", "(+ 2 (* 3 4))"},
		{"2*3+4", "(+ (* 2 3) 4)"},

		// Left-associative + and -
		{"1+2+3", "(+ (+ 1 2) 3)"},
		{"1-2-3", "(- (- 1 2) 3)"},

		// Left-associative * and /
		{"2*3*4", "(* (* 2 3) 4)"},
		{"12/3/2", "(/ (/ 12 3) 2)"},

		// Left-associative ^ (matches expected behavior)
		{"2^3^4", "(^ (^ 2 3) 4)"},

		// Mixed precedence
		{"1+2*3^4", "(+ 1 (* 2 (^ 3 4)))"},

		// Comparison binds loosest (among binary ops)
		{"A1+1>B1*2", "(> (+ A1 1) (* B1 2))"},
		{"A1=B1", "(= A1 B1)"},
		{"A1<>B1", "(<> A1 B1)"},
		{"A1<=B1", "(<= A1 B1)"},
		{"A1>=B1", "(>= A1 B1)"},

		// Concat between comparison and addition
		{`A1&"x"&B1`, `(& (& A1 "x") B1)`},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseUnary(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"-1", "(- 1)"},
		{"+1", "(+ 1)"},
		{"--1", "(- (- 1))"},

		// Unary binds tighter than ^ (convention: -2^2 = 4)
		{"-2^3", "(^ (- 2) 3)"}, // (-2)^3 = -8 (matches expected behavior)
		{"-1*2", "(* (- 1) 2)"}, // (-1)*2 = -2
		{"-1+2", "(+ (- 1) 2)"}, // (-1)+2 = 1
		{"+A1*B1", "(* (+ A1) B1)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParsePostfixPercent(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"50%", "(% 50)"},
		{"50%%", "(% (% 50))"},
		{"A1%", "(% A1)"},
		{"50%*2", "(* (% 50) 2)"},
		{"2*50%", "(* 2 (% 50))"},
		{"(1+2)%", "(% (+ 1 2))"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseCellRefs(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"A1", "A1"},
		{"$A$1", "$A$1"},
		{"A$1", "A$1"},
		{"$A1", "$A1"},
		{"AA100", "AA100"},
		{"Sheet1!A1", "Sheet1!A1"},
		{"Sheet1!$A$1", "Sheet1!$A$1"},
		{"A1+B1", "(+ A1 B1)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseRanges(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"A1:B5", "(: A1 B5)"},
		{"$A$1:$B$5", "(: $A$1 $B$5)"},
		{"Sheet1!A1:B5", "(: Sheet1!A1 B5)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseFullColumnRanges(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// Full-column references are expanded to row 1:1048576.
		{"F:F", "(: F1 F1048576)"},
		{"A:C", "(: A1 C1048576)"},
		{"$A:$A", "(: $A1 $A1048576)"},
		{"Sheet1!A:A", "(: Sheet1!A1 A1048576)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseFullRowRanges(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// Full-row references are expanded to col A:XFD (1:16384).
		{"5:6", "(: A5 XFD6)"},
		{"1:1", "(: A1 XFD1)"},
		{"100:200", "(: A100 XFD200)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseFullRowInFormula(t *testing.T) {
	// The exact type of formula that should work: SUM(5:6)
	input := "SUM(5:6)"
	_, err := Parse(input)
	if err != nil {
		t.Fatalf("Parse(%q) error: %v", input, err)
	}
}

func TestParseFullColumnInFormula(t *testing.T) {
	// The exact formula from problem.xlsx
	input := "MAX(IF(ISNUMBER(Ledger!F:F),ABS(Ledger!F:F)))"
	_, err := Parse(input)
	if err != nil {
		t.Fatalf("Parse(%q) error: %v", input, err)
	}
}

func TestParseGrouping(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"(1+2)*3", "(* (+ 1 2) 3)"},
		{"(1+2)*(3+4)", "(* (+ 1 2) (+ 3 4))"},
		{"((1+2))", "(+ 1 2)"},
		{"(A1)", "A1"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseFunctions(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// Zero-arg.
		{"NOW()", "(NOW)"},
		{"TODAY()", "(TODAY)"},

		// Single arg.
		{"ABS(-1)", "(ABS (- 1))"},

		// Multiple args.
		{"SUM(1,2)", "(SUM 1 2)"},
		{"SUM(1,2,3)", "(SUM 1 2 3)"},
		{"IF(TRUE,1,0)", "(IF TRUE 1 0)"},

		// Range arg.
		{"SUM(A1:A10)", "(SUM (: A1 A10))"},
		{"AVERAGE(B1:B100)", "(AVERAGE (: B1 B100))"},

		// Nested functions.
		{"SUM(IF(A1>0,A1,0),B1)", "(SUM (IF (> A1 0) A1 0) B1)"},
		{"IF(AND(A1>0,B1<100),1,0)", "(IF (AND (> A1 0) (< B1 100)) 1 0)"},

		// INDEX/MATCH.
		{"INDEX(B1:B10,MATCH(D1,A1:A10,0))", "(INDEX (: B1 B10) (MATCH D1 (: A1 A10) 0))"},

		// VLOOKUP with sheet-qualified range.
		{"VLOOKUP(A1,Sheet1!A1:C100,3,FALSE)", "(VLOOKUP A1 (: Sheet1!A1 C100) 3 FALSE)"},

		// IFERROR.
		{"IFERROR(A1/B1,0)", "(IFERROR (/ A1 B1) 0)"},
		{"IFERROR(A1,#N/A)", "(IFERROR A1 #N/A)"},

		// CONCATENATE with string args.
		{`CONCATENATE(A1," ",B1)`, `(CONCATENATE A1 " " B1)`},

		// Function with expression args.
		{"MAX(A1*2,B1+3)", "(MAX (* A1 2) (+ B1 3))"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseArrayLiterals(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// Single row.
		{"{1,2,3}", "{1,2,3}"},

		// Multiple rows.
		{"{1,2;3,4}", "{1,2;3,4}"},

		// String array.
		{`{"a","b";"c","d"}`, `{"a","b";"c","d"}`},

		// Mixed types.
		{"{1,TRUE;#N/A,0}", "{1,TRUE;#N/A,0}"},

		// Single element.
		{"{42}", "{42}"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseComplexFormulas(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// IF(AND(...)) with arithmetic.
		{
			`IF(AND(A1>0,B1<100),A1*B1/100,"N/A")`,
			`(IF (AND (> A1 0) (< B1 100)) (/ (* A1 B1) 100) "N/A")`,
		},

		// SUMPRODUCT with array multiplication.
		{
			"SUMPRODUCT((A1:A10>0)*(B1:B10))",
			"(SUMPRODUCT (* (> (: A1 A10) 0) (: B1 B10)))",
		},

		// Percentage in formula.
		{
			"A1*10%+B1*20%",
			"(+ (* A1 (% 10)) (* B1 (% 20)))",
		},

		// Concat with spaces.
		{
			`A1&" "&B1`,
			`(& (& A1 " ") B1)`,
		},

		// Nested IF.
		{
			"IF(A1>90,\"A\",IF(A1>80,\"B\",\"C\"))",
			`(IF (> A1 90) "A" (IF (> A1 80) "B" "C"))`,
		},

		// SUMIF with string criteria.
		{
			`SUMIF(A1:A10,">0",B1:B10)`,
			`(SUMIF (: A1 A10) ">0" (: B1 B10))`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseErrors(t *testing.T) {
	tests := []struct {
		name  string
		input string
	}{
		{"empty", ""},
		{"trailing operator", "1+"},
		{"unmatched open paren", "(1+2"},
		{"unmatched close paren", ")"},
		{"unexpected operator", "*1"},
		{"two values no operator", "1 2"},
		{"range with non-cellref right", "A1:5"},
		{"incomplete function", "SUM(1,"},
		{"unmatched function paren", "SUM(1,2"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := Parse(tt.input)
			if err == nil {
				t.Errorf("Parse(%q) expected error, got nil", tt.input)
			}
		})
	}
}

func TestParseWhitespace(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{" 1 + 2 ", "(+ 1 2)"},
		{"  SUM( A1 : A10 )  ", "(SUM (: A1 A10))"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			node, err := Parse(tt.input)
			if err != nil {
				t.Fatalf("Parse(%q) error: %v", tt.input, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.input, got, tt.want)
			}
		})
	}
}

func TestParseASTTypes(t *testing.T) {
	// Verify specific AST node types are produced.
	t.Run("NumberLit", func(t *testing.T) {
		node, err := Parse("42")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*NumberLit)
		if !ok {
			t.Fatalf("expected *NumberLit, got %T", node)
		}
		if n.Value != 42 {
			t.Errorf("Value = %f, want 42", n.Value)
		}
	})

	t.Run("StringLit", func(t *testing.T) {
		node, err := Parse(`"hello"`)
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*StringLit)
		if !ok {
			t.Fatalf("expected *StringLit, got %T", node)
		}
		if n.Value != "hello" {
			t.Errorf("Value = %q, want %q", n.Value, "hello")
		}
	})

	t.Run("BoolLit", func(t *testing.T) {
		node, err := Parse("TRUE")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*BoolLit)
		if !ok {
			t.Fatalf("expected *BoolLit, got %T", node)
		}
		if !n.Value {
			t.Error("Value = false, want true")
		}
	})

	t.Run("ErrorLit", func(t *testing.T) {
		node, err := Parse("#DIV/0!")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*ErrorLit)
		if !ok {
			t.Fatalf("expected *ErrorLit, got %T", node)
		}
		if n.Code != ErrDIV0 {
			t.Errorf("Code = %q, want %q", n.Code, ErrDIV0)
		}
	})

	t.Run("CellRef", func(t *testing.T) {
		node, err := Parse("$A$1")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*CellRef)
		if !ok {
			t.Fatalf("expected *CellRef, got %T", node)
		}
		if n.Col != 1 || n.Row != 1 || !n.AbsCol || !n.AbsRow {
			t.Errorf("got %+v", n)
		}
	})

	t.Run("RangeRef", func(t *testing.T) {
		node, err := Parse("A1:B5")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*RangeRef)
		if !ok {
			t.Fatalf("expected *RangeRef, got %T", node)
		}
		if n.From.Col != 1 || n.From.Row != 1 {
			t.Errorf("From = %+v", n.From)
		}
		if n.To.Col != 2 || n.To.Row != 5 {
			t.Errorf("To = %+v", n.To)
		}
	})

	t.Run("FuncCall", func(t *testing.T) {
		node, err := Parse("SUM(1,2)")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*FuncCall)
		if !ok {
			t.Fatalf("expected *FuncCall, got %T", node)
		}
		if n.Name != "SUM" {
			t.Errorf("Name = %q, want %q", n.Name, "SUM")
		}
		if len(n.Args) != 2 {
			t.Errorf("len(Args) = %d, want 2", len(n.Args))
		}
	})

	t.Run("UnaryExpr", func(t *testing.T) {
		node, err := Parse("-1")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*UnaryExpr)
		if !ok {
			t.Fatalf("expected *UnaryExpr, got %T", node)
		}
		if n.Op != "-" {
			t.Errorf("Op = %q, want %q", n.Op, "-")
		}
	})

	t.Run("BinaryExpr", func(t *testing.T) {
		node, err := Parse("1+2")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*BinaryExpr)
		if !ok {
			t.Fatalf("expected *BinaryExpr, got %T", node)
		}
		if n.Op != "+" {
			t.Errorf("Op = %q, want %q", n.Op, "+")
		}
	})

	t.Run("PostfixExpr", func(t *testing.T) {
		node, err := Parse("50%")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*PostfixExpr)
		if !ok {
			t.Fatalf("expected *PostfixExpr, got %T", node)
		}
		if n.Op != "%" {
			t.Errorf("Op = %q, want %q", n.Op, "%")
		}
	})

	t.Run("ArrayLit", func(t *testing.T) {
		node, err := Parse("{1,2;3,4}")
		if err != nil {
			t.Fatal(err)
		}
		n, ok := node.(*ArrayLit)
		if !ok {
			t.Fatalf("expected *ArrayLit, got %T", node)
		}
		if len(n.Rows) != 2 {
			t.Errorf("len(Rows) = %d, want 2", len(n.Rows))
		}
		if len(n.Rows[0]) != 2 || len(n.Rows[1]) != 2 {
			t.Errorf("row lengths = %d, %d, want 2, 2", len(n.Rows[0]), len(n.Rows[1]))
		}
	})
}
