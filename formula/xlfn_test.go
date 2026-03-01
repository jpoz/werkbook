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
		{name: "MAXIFS gets prefix", in: "MAXIFS(A1:A5,B1:B5,1)", want: "_xlfn.MAXIFS(A1:A5,B1:B5,1)"},
		{name: "IFS gets prefix", in: "IFS(A1>0,1,A1<0,-1)", want: "_xlfn.IFS(A1>0,1,A1<0,-1)"},
		{name: "SORT gets xlws prefix", in: "SORT(A1:A5)", want: "_xlfn._xlws.SORT(A1:A5)"},
		{name: "XLOOKUP gets prefix", in: "XLOOKUP(A1,B:B,C:C)", want: "_xlfn.XLOOKUP(A1,B:B,C:C)"},
		{name: "TEXTJOIN gets prefix", in: `TEXTJOIN(",",TRUE,A1:A5)`, want: `_xlfn.TEXTJOIN(",",TRUE,A1:A5)`},
		{name: "nested xlfn functions", in: "MAXIFS(A1:A5,B1:B5,IFS(C1>0,1))", want: "_xlfn.MAXIFS(A1:A5,B1:B5,_xlfn.IFS(C1>0,1))"},
		{name: "mixed legacy and xlfn", in: "SUM(MAXIFS(A1:A5,B1:B5,1))", want: "SUM(_xlfn.MAXIFS(A1:A5,B1:B5,1))"},
		{name: "string containing function name", in: `CONCAT("MAXIFS",A1)`, want: `_xlfn.CONCAT("MAXIFS",A1)`},
		{name: "already prefixed no double", in: "_xlfn.MAXIFS(A1:A5,B1:B5,1)", want: "_xlfn.MAXIFS(A1:A5,B1:B5,1)"},
		{name: "already prefixed xlws no double", in: "_xlfn._xlws.SORT(A1:A5)", want: "_xlfn._xlws.SORT(A1:A5)"},
		{name: "IFERROR unchanged", in: "IFERROR(A1/B1,0)", want: "IFERROR(A1/B1,0)"},
		{name: "XOR", in: "XOR(A1,B1)", want: "_xlfn.XOR(A1,B1)"},
		{name: "SWITCH", in: "SWITCH(A1,1,\"one\",2,\"two\")", want: "_xlfn.SWITCH(A1,1,\"one\",2,\"two\")"},
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
		{name: "strip xlfn", in: "_xlfn.MAXIFS(A1:A5,B1:B5,1)", want: "MAXIFS(A1:A5,B1:B5,1)"},
		{name: "strip xlws", in: "_xlfn._xlws.SORT(A1:A5)", want: "SORT(A1:A5)"},
		{name: "strip nested", in: "_xlfn.MAXIFS(A1:A5,B1:B5,_xlfn.IFS(C1>0,1))", want: "MAXIFS(A1:A5,B1:B5,IFS(C1>0,1))"},
		{name: "strip mixed", in: "SUM(_xlfn.MAXIFS(A1:A5,B1:B5,1))", want: "SUM(MAXIFS(A1:A5,B1:B5,1))"},
		{name: "legacy unchanged", in: "IF(A1>0,1,0)", want: "IF(A1>0,1,0)"},
		{name: "IFERROR", in: "_xlfn.IFERROR(A1/B1,0)", want: "IFERROR(A1/B1,0)"},
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
		"MAXIFS(A1:A5,B1:B5,1)",
		"SORT(A1:A5)",
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
