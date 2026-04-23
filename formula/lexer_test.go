package formula

import (
	"testing"
)

// tok is a shorthand for building expected tokens in tests.
func tok(typ TokenType, val string) Token {
	return Token{Type: typ, Value: val}
}

// tokensEqual compares two token slices ignoring Pos fields.
func tokensEqual(a, b []Token) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i].Type != b[i].Type || a[i].Value != b[i].Value {
			return false
		}
	}
	return true
}

func TestTokenizeSimpleArithmetic(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "1+2",
			want: []Token{
				tok(TokNumber, "1"),
				tok(TokOp, "+"),
				tok(TokNumber, "2"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "1 + 2",
			want: []Token{
				tok(TokNumber, "1"),
				tok(TokOp, "+"),
				tok(TokNumber, "2"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "10*20+3/4-1",
			want: []Token{
				tok(TokNumber, "10"),
				tok(TokOp, "*"),
				tok(TokNumber, "20"),
				tok(TokOp, "+"),
				tok(TokNumber, "3"),
				tok(TokOp, "/"),
				tok(TokNumber, "4"),
				tok(TokOp, "-"),
				tok(TokNumber, "1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "2^3",
			want: []Token{
				tok(TokNumber, "2"),
				tok(TokOp, "^"),
				tok(TokNumber, "3"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeNumbers(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "123",
			want:  []Token{tok(TokNumber, "123"), tok(TokEOF, "")},
		},
		{
			input: "1.5",
			want:  []Token{tok(TokNumber, "1.5"), tok(TokEOF, "")},
		},
		{
			input: ".5",
			want:  []Token{tok(TokNumber, ".5"), tok(TokEOF, "")},
		},
		{
			input: "1.5E10",
			want:  []Token{tok(TokNumber, "1.5E10"), tok(TokEOF, "")},
		},
		{
			input: "1.5e-3",
			want:  []Token{tok(TokNumber, "1.5e-3"), tok(TokEOF, "")},
		},
		{
			input: "2E+5",
			want:  []Token{tok(TokNumber, "2E+5"), tok(TokEOF, "")},
		},
		{
			input: "0",
			want:  []Token{tok(TokNumber, "0"), tok(TokEOF, "")},
		},
		{
			input: "100.00",
			want:  []Token{tok(TokNumber, "100.00"), tok(TokEOF, "")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeStrings(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: `"hello"`,
			want:  []Token{tok(TokString, "hello"), tok(TokEOF, "")},
		},
		{
			input: `""`,
			want:  []Token{tok(TokString, ""), tok(TokEOF, "")},
		},
		{
			input: `"he said ""hi"""`,
			want:  []Token{tok(TokString, `he said "hi"`), tok(TokEOF, "")},
		},
		{
			input: `"line1" & "line2"`,
			want: []Token{
				tok(TokString, "line1"),
				tok(TokOp, "&"),
				tok(TokString, "line2"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeStringUnterminated(t *testing.T) {
	_, err := Tokenize(`"hello`)
	if err == nil {
		t.Fatal("expected error for unterminated string")
	}
}

func TestTokenizeBooleans(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "TRUE",
			want:  []Token{tok(TokBool, "TRUE"), tok(TokEOF, "")},
		},
		{
			input: "FALSE",
			want:  []Token{tok(TokBool, "FALSE"), tok(TokEOF, "")},
		},
		{
			input: "true",
			want:  []Token{tok(TokBool, "TRUE"), tok(TokEOF, "")},
		},
		{
			input: "True",
			want:  []Token{tok(TokBool, "TRUE"), tok(TokEOF, "")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeErrors(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "#N/A",
			want:  []Token{tok(TokError, "#N/A"), tok(TokEOF, "")},
		},
		{
			input: "#DIV/0!",
			want:  []Token{tok(TokError, "#DIV/0!"), tok(TokEOF, "")},
		},
		{
			input: "#VALUE!",
			want:  []Token{tok(TokError, "#VALUE!"), tok(TokEOF, "")},
		},
		{
			input: "#REF!",
			want:  []Token{tok(TokError, "#REF!"), tok(TokEOF, "")},
		},
		{
			input: "#NAME?",
			want:  []Token{tok(TokError, "#NAME?"), tok(TokEOF, "")},
		},
		{
			input: "#NUM!",
			want:  []Token{tok(TokError, "#NUM!"), tok(TokEOF, "")},
		},
		{
			input: "#NULL!",
			want:  []Token{tok(TokError, "#NULL!"), tok(TokEOF, "")},
		},
		{
			input: "#SPILL!",
			want:  []Token{tok(TokError, "#SPILL!"), tok(TokEOF, "")},
		},
		{
			input: "#CALC!",
			want:  []Token{tok(TokError, "#CALC!"), tok(TokEOF, "")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeCellRefs(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "A1",
			want:  []Token{tok(TokCellRef, "A1"), tok(TokEOF, "")},
		},
		{
			input: "$A$1",
			want:  []Token{tok(TokCellRef, "$A$1"), tok(TokEOF, "")},
		},
		{
			input: "A$1",
			want:  []Token{tok(TokCellRef, "A$1"), tok(TokEOF, "")},
		},
		{
			input: "$A1",
			want:  []Token{tok(TokCellRef, "$A1"), tok(TokEOF, "")},
		},
		{
			input: "XFD1048576",
			want:  []Token{tok(TokCellRef, "XFD1048576"), tok(TokEOF, "")},
		},
		{
			input: "AA100",
			want:  []Token{tok(TokCellRef, "AA100"), tok(TokEOF, "")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeRanges(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "A1:B5",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokColon, ":"),
				tok(TokCellRef, "B5"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "$A$1:$B$5",
			want: []Token{
				tok(TokCellRef, "$A$1"),
				tok(TokColon, ":"),
				tok(TokCellRef, "$B$5"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeSheetQualifiedRefs(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "Sheet1!A1",
			want:  []Token{tok(TokCellRef, "Sheet1!A1"), tok(TokEOF, "")},
		},
		{
			input: "'Sheet Name'!A1",
			want:  []Token{tok(TokCellRef, "'Sheet Name'!A1"), tok(TokEOF, "")},
		},
		{
			input: "'Sheet Name'!$A$1",
			want:  []Token{tok(TokCellRef, "'Sheet Name'!$A$1"), tok(TokEOF, "")},
		},
		{
			input: "'It''s a sheet'!B2",
			want:  []Token{tok(TokCellRef, "'It''s a sheet'!B2"), tok(TokEOF, "")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeFunctions(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "SUM(A1:A10)",
			want: []Token{
				tok(TokFunc, "SUM("),
				tok(TokCellRef, "A1"),
				tok(TokColon, ":"),
				tok(TokCellRef, "A10"),
				tok(TokRParen, ")"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "IF(A1>0,A1,-A1)",
			want: []Token{
				tok(TokFunc, "IF("),
				tok(TokCellRef, "A1"),
				tok(TokOp, ">"),
				tok(TokNumber, "0"),
				tok(TokComma, ","),
				tok(TokCellRef, "A1"),
				tok(TokComma, ","),
				tok(TokOp, "-"),
				tok(TokCellRef, "A1"),
				tok(TokRParen, ")"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "AVERAGE(B1:B100)",
			want: []Token{
				tok(TokFunc, "AVERAGE("),
				tok(TokCellRef, "B1"),
				tok(TokColon, ":"),
				tok(TokCellRef, "B100"),
				tok(TokRParen, ")"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeNestedFunctions(t *testing.T) {
	input := "SUM(IF(A1>0,A1,0),B1)"
	want := []Token{
		tok(TokFunc, "SUM("),
		tok(TokFunc, "IF("),
		tok(TokCellRef, "A1"),
		tok(TokOp, ">"),
		tok(TokNumber, "0"),
		tok(TokComma, ","),
		tok(TokCellRef, "A1"),
		tok(TokComma, ","),
		tok(TokNumber, "0"),
		tok(TokRParen, ")"),
		tok(TokComma, ","),
		tok(TokCellRef, "B1"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeComparisonOperators(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "A1=B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, "="),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A1<>B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, "<>"),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A1<=B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, "<="),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A1>=B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, ">="),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A1<B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, "<"),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A1>B1",
			want: []Token{
				tok(TokCellRef, "A1"),
				tok(TokOp, ">"),
				tok(TokCellRef, "B1"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizePercent(t *testing.T) {
	input := "50%"
	want := []Token{
		tok(TokNumber, "50"),
		tok(TokPercent, "%"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeConcat(t *testing.T) {
	input := `A1&" "&B1`
	want := []Token{
		tok(TokCellRef, "A1"),
		tok(TokOp, "&"),
		tok(TokString, " "),
		tok(TokOp, "&"),
		tok(TokCellRef, "B1"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeArrayLiteral(t *testing.T) {
	input := "{1,2;3,4}"
	want := []Token{
		tok(TokArrayOpen, "{"),
		tok(TokNumber, "1"),
		tok(TokComma, ","),
		tok(TokNumber, "2"),
		tok(TokSemicolon, ";"),
		tok(TokNumber, "3"),
		tok(TokComma, ","),
		tok(TokNumber, "4"),
		tok(TokArrayClose, "}"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeParenthesizedExpression(t *testing.T) {
	input := "(1+2)*3"
	want := []Token{
		tok(TokLParen, "("),
		tok(TokNumber, "1"),
		tok(TokOp, "+"),
		tok(TokNumber, "2"),
		tok(TokRParen, ")"),
		tok(TokOp, "*"),
		tok(TokNumber, "3"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeComplexFormula(t *testing.T) {
	// VLOOKUP(A1,Sheet2!$A$1:$C$100,3,FALSE)
	input := "VLOOKUP(A1,Sheet2!$A$1:$C$100,3,FALSE)"
	want := []Token{
		tok(TokFunc, "VLOOKUP("),
		tok(TokCellRef, "A1"),
		tok(TokComma, ","),
		tok(TokCellRef, "Sheet2!$A$1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "$C$100"),
		tok(TokComma, ","),
		tok(TokNumber, "3"),
		tok(TokComma, ","),
		tok(TokBool, "FALSE"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeSUMIFFormula(t *testing.T) {
	input := `SUMIF(A1:A10,">0",B1:B10)`
	want := []Token{
		tok(TokFunc, "SUMIF("),
		tok(TokCellRef, "A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "A10"),
		tok(TokComma, ","),
		tok(TokString, ">0"),
		tok(TokComma, ","),
		tok(TokCellRef, "B1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "B10"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeIFERRORFormula(t *testing.T) {
	input := "IFERROR(A1/B1,0)"
	want := []Token{
		tok(TokFunc, "IFERROR("),
		tok(TokCellRef, "A1"),
		tok(TokOp, "/"),
		tok(TokCellRef, "B1"),
		tok(TokComma, ","),
		tok(TokNumber, "0"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeSheetRangeFormula(t *testing.T) {
	input := "SUM('Q1 Data'!A1:A100)"
	want := []Token{
		tok(TokFunc, "SUM("),
		tok(TokCellRef, "'Q1 Data'!A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "A100"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeWhitespaceHandling(t *testing.T) {
	input := "  SUM( A1 : A10 )  "
	want := []Token{
		tok(TokFunc, "SUM("),
		tok(TokCellRef, "A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "A10"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeEmptyFormula(t *testing.T) {
	got, err := Tokenize("")
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", "", err)
	}
	if len(got) != 1 || got[0].Type != TokEOF {
		t.Errorf("expected single EOF token, got %v", got)
	}
}

func TestTokenizePositions(t *testing.T) {
	// Verify that Pos is tracked correctly.
	input := "A1+B2"
	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	// A1 at 0, + at 2, B2 at 3, EOF at 5
	positions := []int{0, 2, 3, 5}
	for i, want := range positions {
		if got[i].Pos != want {
			t.Errorf("token %d (%v): Pos = %d, want %d", i, got[i], got[i].Pos, want)
		}
	}
}

func TestTokenizeErrorInFormula(t *testing.T) {
	// IFERROR with an error literal
	input := "IFERROR(A1,#N/A)"
	want := []Token{
		tok(TokFunc, "IFERROR("),
		tok(TokCellRef, "A1"),
		tok(TokComma, ","),
		tok(TokError, "#N/A"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeMultiplePercentage(t *testing.T) {
	input := "A1*10%+B1*20%"
	want := []Token{
		tok(TokCellRef, "A1"),
		tok(TokOp, "*"),
		tok(TokNumber, "10"),
		tok(TokPercent, "%"),
		tok(TokOp, "+"),
		tok(TokCellRef, "B1"),
		tok(TokOp, "*"),
		tok(TokNumber, "20"),
		tok(TokPercent, "%"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeNegativeNumber(t *testing.T) {
	// Note: the lexer emits - as an operator; the parser handles unary.
	input := "-1+2"
	want := []Token{
		tok(TokOp, "-"),
		tok(TokNumber, "1"),
		tok(TokOp, "+"),
		tok(TokNumber, "2"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeStringArray(t *testing.T) {
	input := `{"a","b";"c","d"}`
	want := []Token{
		tok(TokArrayOpen, "{"),
		tok(TokString, "a"),
		tok(TokComma, ","),
		tok(TokString, "b"),
		tok(TokSemicolon, ";"),
		tok(TokString, "c"),
		tok(TokComma, ","),
		tok(TokString, "d"),
		tok(TokArrayClose, "}"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeINDEXMATCH(t *testing.T) {
	input := "INDEX(B1:B10,MATCH(D1,A1:A10,0))"
	want := []Token{
		tok(TokFunc, "INDEX("),
		tok(TokCellRef, "B1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "B10"),
		tok(TokComma, ","),
		tok(TokFunc, "MATCH("),
		tok(TokCellRef, "D1"),
		tok(TokComma, ","),
		tok(TokCellRef, "A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "A10"),
		tok(TokComma, ","),
		tok(TokNumber, "0"),
		tok(TokRParen, ")"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeLooksLikeCellRef(t *testing.T) {
	tests := []struct {
		input string
		want  bool
	}{
		{"A1", true},
		{"$A$1", true},
		{"A$1", true},
		{"$A1", true},
		{"XFD1048576", true},
		{"AA100", true},
		{"AAAA1", false}, // 4 letters
		{"A", false},     // no digits
		{"1A", false},    // starts with digit
		{"$1", false},    // no letters
		{"", false},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := looksLikeCellRef(tt.input)
			if got != tt.want {
				t.Errorf("looksLikeCellRef(%q) = %v, want %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeSUMPRODUCT(t *testing.T) {
	input := "SUMPRODUCT((A1:A10>0)*(B1:B10))"
	want := []Token{
		tok(TokFunc, "SUMPRODUCT("),
		tok(TokLParen, "("),
		tok(TokCellRef, "A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "A10"),
		tok(TokOp, ">"),
		tok(TokNumber, "0"),
		tok(TokRParen, ")"),
		tok(TokOp, "*"),
		tok(TokLParen, "("),
		tok(TokCellRef, "B1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "B10"),
		tok(TokRParen, ")"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeMultiSheetRange(t *testing.T) {
	input := "Sheet1!A1:Sheet1!A10"
	want := []Token{
		tok(TokCellRef, "Sheet1!A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "Sheet1!A10"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeComplexNestedFormula(t *testing.T) {
	// A realistic complex formula
	input := `IF(AND(A1>0,B1<100),A1*B1/100,"N/A")`
	want := []Token{
		tok(TokFunc, "IF("),
		tok(TokFunc, "AND("),
		tok(TokCellRef, "A1"),
		tok(TokOp, ">"),
		tok(TokNumber, "0"),
		tok(TokComma, ","),
		tok(TokCellRef, "B1"),
		tok(TokOp, "<"),
		tok(TokNumber, "100"),
		tok(TokRParen, ")"),
		tok(TokComma, ","),
		tok(TokCellRef, "A1"),
		tok(TokOp, "*"),
		tok(TokCellRef, "B1"),
		tok(TokOp, "/"),
		tok(TokNumber, "100"),
		tok(TokComma, ","),
		tok(TokString, "N/A"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeUnknownErrorLiteral(t *testing.T) {
	_, err := Tokenize("#BOGUS!")
	if err == nil {
		t.Fatal("expected error for unknown error literal")
	}
}

func TestTokenizeUnexpectedChar(t *testing.T) {
	_, err := Tokenize("A1 ¤ B1")
	if err == nil {
		t.Fatal("expected error for unexpected character")
	}
}

func TestTokenizeAtPrefix(t *testing.T) {
	got, err := Tokenize("@A1")
	if err != nil {
		t.Fatalf("Tokenize(@A1) error: %v", err)
	}
	want := []Token{
		tok(TokAt, "@"),
		tok(TokCellRef, "A1"),
		tok(TokEOF, ""),
	}
	if !tokensEqual(got, want) {
		t.Fatalf("Tokenize(@A1)\n  got:  %v\n  want: %v", got, want)
	}
}

func TestTokenizeBoolInExpression(t *testing.T) {
	input := "IF(TRUE,1,0)"
	want := []Token{
		tok(TokFunc, "IF("),
		tok(TokBool, "TRUE"),
		tok(TokComma, ","),
		tok(TokNumber, "1"),
		tok(TokComma, ","),
		tok(TokNumber, "0"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeQuotedSheetRange(t *testing.T) {
	input := "'My Sheet'!A1:'My Sheet'!B10"
	want := []Token{
		tok(TokCellRef, "'My Sheet'!A1"),
		tok(TokColon, ":"),
		tok(TokCellRef, "'My Sheet'!B10"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}

func TestTokenizeFullColumnRef(t *testing.T) {
	tests := []struct {
		input string
		want  []Token
	}{
		{
			input: "F:F",
			want: []Token{
				tok(TokCellRef, "F"),
				tok(TokColon, ":"),
				tok(TokCellRef, "F"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "$A:$A",
			want: []Token{
				tok(TokCellRef, "$A"),
				tok(TokColon, ":"),
				tok(TokCellRef, "$A"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "A:C",
			want: []Token{
				tok(TokCellRef, "A"),
				tok(TokColon, ":"),
				tok(TokCellRef, "C"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "Sheet1!A:A",
			want: []Token{
				tok(TokCellRef, "Sheet1!A"),
				tok(TokColon, ":"),
				tok(TokCellRef, "A"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "Ledger!F:F",
			want: []Token{
				tok(TokCellRef, "Ledger!F"),
				tok(TokColon, ":"),
				tok(TokCellRef, "F"),
				tok(TokEOF, ""),
			},
		},
		{
			input: "'Sheet Name'!B:B",
			want: []Token{
				tok(TokCellRef, "'Sheet Name'!B"),
				tok(TokColon, ":"),
				tok(TokCellRef, "B"),
				tok(TokEOF, ""),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := Tokenize(tt.input)
			if err != nil {
				t.Fatalf("Tokenize(%q) error: %v", tt.input, err)
			}
			if !tokensEqual(got, tt.want) {
				t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", tt.input, got, tt.want)
			}
		})
	}
}

func TestTokenizeCONCATENATE(t *testing.T) {
	input := `CONCATENATE(A1," ",B1," ",C1)`
	want := []Token{
		tok(TokFunc, "CONCATENATE("),
		tok(TokCellRef, "A1"),
		tok(TokComma, ","),
		tok(TokString, " "),
		tok(TokComma, ","),
		tok(TokCellRef, "B1"),
		tok(TokComma, ","),
		tok(TokString, " "),
		tok(TokComma, ","),
		tok(TokCellRef, "C1"),
		tok(TokRParen, ")"),
		tok(TokEOF, ""),
	}

	got, err := Tokenize(input)
	if err != nil {
		t.Fatalf("Tokenize(%q) error: %v", input, err)
	}
	if !tokensEqual(got, want) {
		t.Errorf("Tokenize(%q)\n  got:  %v\n  want: %v", input, got, want)
	}
}
