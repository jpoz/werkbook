package formula

import "testing"

func TestAddXlfnPrefixes(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{name: "empty", in: "", want: ""},
		{name: "no functions", in: "A1+B2", want: "A1+B2"},
		{name: "legacy SUM unchanged", in: "SUM(A1:A5)", want: "SUM(A1:A5)"},
		{name: "legacy IF unchanged", in: "IF(A1>0,1,0)", want: "IF(A1>0,1,0)"},
		{name: "FILTER gets xlws prefix", in: `FILTER(A1:A5,A1:A5<>"")`, want: `_xlfn._xlws.FILTER(A1:A5,A1:A5<>"")`},
		{name: "MAXIFS gets prefix", in: "MAXIFS(A1:A5,B1:B5,1)", want: "_xlfn.MAXIFS(A1:A5,B1:B5,1)"},
		{name: "IFS gets prefix", in: "IFS(A1>0,1,A1<0,-1)", want: "_xlfn.IFS(A1>0,1,A1<0,-1)"},
		{name: "SORT gets xlws prefix", in: "SORT(A1:A5)", want: "_xlfn._xlws.SORT(A1:A5)"},
		{name: "XLOOKUP gets prefix", in: "XLOOKUP(A1,B:B,C:C)", want: "_xlfn.XLOOKUP(A1,B:B,C:C)"},
		{name: "TEXTJOIN gets prefix", in: `TEXTJOIN(",",TRUE,A1:A5)`, want: `_xlfn.TEXTJOIN(",",TRUE,A1:A5)`},
		{name: "UNIQUE gets prefix", in: "UNIQUE(A1:A5)", want: "_xlfn.UNIQUE(A1:A5)"},
		{name: "SINGLE gets prefix", in: "SINGLE(A1)", want: "_xlfn.SINGLE(A1)"},
		{name: "nested xlfn functions", in: "MAXIFS(A1:A5,B1:B5,IFS(C1>0,1))", want: "_xlfn.MAXIFS(A1:A5,B1:B5,_xlfn.IFS(C1>0,1))"},
		{name: "nested dynamic array functions", in: `SORT(UNIQUE(FILTER(A1:A10,A1:A10<>"")))`, want: `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(A1:A10,A1:A10<>"")))`},
		{name: "mixed legacy and xlfn", in: "SUM(MAXIFS(A1:A5,B1:B5,1))", want: "SUM(_xlfn.MAXIFS(A1:A5,B1:B5,1))"},
		{name: "string containing function name", in: `CONCAT("MAXIFS",A1)`, want: `_xlfn.CONCAT("MAXIFS",A1)`},
		{name: "already prefixed no double", in: "_xlfn.MAXIFS(A1:A5,B1:B5,1)", want: "_xlfn.MAXIFS(A1:A5,B1:B5,1)"},
		{name: "already prefixed FILTER no double", in: `_xlfn._xlws.FILTER(A1:A5,A1:A5<>"")`, want: `_xlfn._xlws.FILTER(A1:A5,A1:A5<>"")`},
		{name: "already prefixed xlws no double", in: "_xlfn._xlws.SORT(A1:A5)", want: "_xlfn._xlws.SORT(A1:A5)"},
		{name: "IFERROR unchanged", in: "IFERROR(A1/B1,0)", want: "IFERROR(A1/B1,0)"},
		{name: "XOR", in: "XOR(A1,B1)", want: "_xlfn.XOR(A1,B1)"},
		{name: "SWITCH", in: "SWITCH(A1,1,\"one\",2,\"two\")", want: "_xlfn.SWITCH(A1,1,\"one\",2,\"two\")"},
		{name: "LET", in: "LET(x,5,x+1)", want: "_xlfn.LET(_xlpm.x,5,_xlpm.x+1)"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := AddXlfnPrefixes(tt.in)
			if got != tt.want {
				t.Errorf("AddXlfnPrefixes(%q) = %q, want %q", tt.in, got, tt.want)
			}
		})
	}
}

func TestStripXlfnPrefixes(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{name: "empty", in: "", want: ""},
		{name: "no prefix", in: "SUM(A1:A5)", want: "SUM(A1:A5)"},
		{name: "strip FILTER xlws", in: `_xlfn._xlws.FILTER(A1:A5,A1:A5<>"")`, want: `FILTER(A1:A5,A1:A5<>"")`},
		{name: "strip xlfn", in: "_xlfn.MAXIFS(A1:A5,B1:B5,1)", want: "MAXIFS(A1:A5,B1:B5,1)"},
		{name: "strip xlws", in: "_xlfn._xlws.SORT(A1:A5)", want: "SORT(A1:A5)"},
		{name: "strip UNIQUE", in: "_xlfn.UNIQUE(A1:A5)", want: "UNIQUE(A1:A5)"},
		{name: "strip nested", in: "_xlfn.MAXIFS(A1:A5,B1:B5,_xlfn.IFS(C1>0,1))", want: "MAXIFS(A1:A5,B1:B5,IFS(C1>0,1))"},
		{name: "strip nested dynamic array functions", in: `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(A1:A10,A1:A10<>"")))`, want: `SORT(UNIQUE(FILTER(A1:A10,A1:A10<>"")))`},
		{name: "strip mixed", in: "SUM(_xlfn.MAXIFS(A1:A5,B1:B5,1))", want: "SUM(MAXIFS(A1:A5,B1:B5,1))"},
		{name: "legacy unchanged", in: "IF(A1>0,1,0)", want: "IF(A1>0,1,0)"},
		{name: "IFERROR", in: "_xlfn.IFERROR(A1/B1,0)", want: "IFERROR(A1/B1,0)"},
		{name: "LET", in: "_xlfn.LET(_xlpm.x,5,_xlpm.x+1)", want: "LET(x,5,x+1)"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := StripXlfnPrefixes(tt.in)
			if got != tt.want {
				t.Errorf("StripXlfnPrefixes(%q) = %q, want %q", tt.in, got, tt.want)
			}
		})
	}
}

func TestAddStripRoundTrip(t *testing.T) {
	formulas := []string{
		`FILTER(A1:A5,A1:A5<>"")`,
		"MAXIFS(A1:A5,B1:B5,1)",
		"SORT(A1:A5)",
		"UNIQUE(A1:A5)",
		"SINGLE(A1)",
		`SORT(UNIQUE(FILTER(A1:A10,A1:A10<>"")))`,
		"SUM(MAXIFS(A1:A5,B1:B5,IFS(C1>0,1)))",
		"SUM(A1:A5)",
		"IF(A1>0,1,0)",
	}

	for _, f := range formulas {
		t.Run(f, func(t *testing.T) {
			added := AddXlfnPrefixes(f)
			stripped := StripXlfnPrefixes(added)
			if stripped != f {
				t.Errorf("round-trip failed: %q -> %q -> %q", f, added, stripped)
			}
		})
	}
}

func TestIsDynamicArrayFormula(t *testing.T) {
	tests := []struct {
		formula string
		want    bool
	}{
		{formula: "", want: false},
		{formula: "SUM(A1:A5)", want: false},
		{formula: "FILTER(A1:A5,A1:A5<>\"\")", want: true},
		{formula: "_xlfn.ANCHORARRAY(A1)", want: true},
		{formula: "_xlfn.SINGLE(A1)", want: true},
		{formula: "XLOOKUP(A1,B:B,C:C)", want: true},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			if got := IsDynamicArrayFormula(tt.formula); got != tt.want {
				t.Fatalf("IsDynamicArrayFormula(%q) = %v, want %v", tt.formula, got, tt.want)
			}
		})
	}
}
