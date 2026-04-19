package formula

import "testing"

func TestImmediateLambdaInvocation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Data", Col: 1, Row: 1}: NumberVal(10),
		},
	}

	tests := []struct {
		formula string
		want    Value
	}{
		{`LAMBDA(x, x+1)(5)`, NumberVal(6)},
		{`LAMBDA(x,y, x+y)(3, 4)`, NumberVal(7)},
		{`LAMBDA(a,b, a&b)("hello", " world")`, StringVal("hello world")},
		{`LAMBDA(42)()`, NumberVal(42)},
		{`LAMBDA(x, x*2)(Data!A1)`, NumberVal(20)},
		{`LAMBDA(x, LAMBDA(y, x+y)(10))(5)`, NumberVal(15)},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.Type {
				t.Fatalf("Eval(%q) type = %v, want %v", tt.formula, got.Type, tt.want.Type)
			}
			switch got.Type {
			case ValueNumber:
				if got.Num != tt.want.Num {
					t.Fatalf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.Num)
				}
			case ValueString:
				if got.Str != tt.want.Str {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.Str)
				}
			default:
				t.Fatalf("unexpected test value type %v", got.Type)
			}
		})
	}
}

func TestISOMITTED(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    Value
	}{
		// ISOMITTED returns TRUE for an omitted lambda parameter.
		{`LET(addOpt, LAMBDA(a,b, IF(ISOMITTED(b), a+10, a+b)), addOpt(5, 3))`, NumberVal(8)},
		{`LET(addOpt, LAMBDA(a,b, IF(ISOMITTED(b), a+10, a+b)), addOpt(5,))`, NumberVal(15)},
		// Omitted middle argument.
		{`LET(greet, LAMBDA(first,mid,last, IF(ISOMITTED(mid), first&" "&last, first&" "&mid&" "&last)), greet("John",, "Doe"))`, StringVal("John Doe")},
		{`LET(greet, LAMBDA(first,mid,last, IF(ISOMITTED(mid), first&" "&last, first&" "&mid&" "&last)), greet("John", "Q", "Doe"))`, StringVal("John Q Doe")},
		// Immediate invocation.
		{`LAMBDA(x, IF(ISOMITTED(x), -1, x*2))(5)`, NumberVal(10)},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.Type {
				t.Fatalf("Eval(%q) type = %v, want %v", tt.formula, got.Type, tt.want.Type)
			}
			switch got.Type {
			case ValueNumber:
				if got.Num != tt.want.Num {
					t.Fatalf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.Num)
				}
			case ValueString:
				if got.Str != tt.want.Str {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.Str)
				}
			}
		})
	}
}

func TestImmediateLambdaInvocationErrors(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFERROR(LAMBDA(x,y,x/y)(1,0), "err")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Fatalf(`IFERROR(LAMBDA(x,y,x/y)(1,0), "err") = %v, want "err"`, got)
	}
}
