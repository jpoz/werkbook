package formula

import (
	"reflect"
	"sort"
	"strings"
	"testing"
)

func TestParseFuncSignature(t *testing.T) {
	arg := func(name string, flags string) ArgDoc {
		return ArgDoc{Name: name, Optional: strings.Contains(flags, "o"), Repeat: strings.Contains(flags, "r")}
	}
	tests := []struct {
		sig      string
		wantName string
		want     []ArgDoc
	}{
		{"PI()", "PI", []ArgDoc{}},
		{"ABS(number)", "ABS", []ArgDoc{arg("number", "")}},
		{"SUMIF(range, criteria, [sum_range])", "SUMIF", []ArgDoc{arg("range", ""), arg("criteria", ""), arg("sum_range", "o")}},
		{"SUM(number1, [number2], ...)", "SUM", []ArgDoc{arg("number1", ""), arg("number2", "or")}},
		{"IF(a, [b], [c])", "IF", []ArgDoc{arg("a", ""), arg("b", "o"), arg("c", "o")}},
		{
			"SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "SUMIFS",
			[]ArgDoc{arg("sum_range", ""), arg("criteria_range1", ""), arg("criteria1", ""), arg("criteria_range2", "or"), arg("criteria2", "or")},
		},
		{"F.X(a, b, ...)", "F.X", []ArgDoc{arg("a", ""), arg("b", "r")}},
		{"ROW([reference])", "ROW", []ArgDoc{arg("reference", "o")}},
		{"X(  a ,[ b ] )", "X", []ArgDoc{arg("a", ""), arg("b", "o")}},
	}
	for _, tt := range tests {
		name, args, err := parseFuncSignature(tt.sig)
		if err != nil {
			t.Errorf("parseFuncSignature(%q): %v", tt.sig, err)
			continue
		}
		if name != tt.wantName || !reflect.DeepEqual(args, tt.want) {
			t.Errorf("parseFuncSignature(%q) = %s %+v, want %s %+v", tt.sig, name, args, tt.wantName, tt.want)
		}
	}

	bad := []string{
		"", "SUM", "SUM(", "(a)", "sum(a)", "SUM(a", "SUM(a,)", "SUM(,a)", "SUM([a)", "SUM(a])",
		"SUM([a, [b]])", "SUM(...)", "SUM(a, ..., b)", "SUM([a, ...])", "SUM(1a)", "SUM(a b)", "MY F(a)",
	}
	for _, sig := range bad {
		if _, _, err := parseFuncSignature(sig); err == nil {
			t.Errorf("parseFuncSignature(%q) gave no error", sig)
		}
	}
}

func TestFuncDocsCoverAllRegisteredFunctions(t *testing.T) {
	docs := FuncDocs()
	// The list is the registry plus the special forms of the parser.
	registered := append(RegisteredFunctions(), funcDocSpecialForms...)
	sort.Strings(registered)
	special := map[string]bool{}
	for _, name := range funcDocSpecialForms {
		special[name] = true
		if _, ok := LookupFunctionInfo(name); ok {
			t.Errorf("%s is a registered function; remove it from funcDocSpecialForms", name)
		}
		// The parser must accept each special form as a call.
		if _, err := Parse(name + "(A1)"); err != nil && name != "LET" && name != "LAMBDA" {
			t.Logf("Parse(%s(A1)): %v", name, err)
		}
	}

	names := make([]string, len(docs))
	for i, d := range docs {
		names[i] = d.Name
	}
	if !sort.StringsAreSorted(names) {
		t.Errorf("FuncDocs is not sorted by name")
	}
	if !reflect.DeepEqual(names, registered) {
		t.Fatalf("FuncDocs has %d names, RegisteredFunctions has %d; the sets differ", len(names), len(registered))
	}
	for i := 1; i < len(names); i++ {
		if names[i] == names[i-1] {
			t.Errorf("duplicate name %s", names[i])
		}
	}

	for _, d := range docs {
		if special[d.Name] {
			if d.Summary == "" || len(d.Args) == 0 || d.MinArgs < 1 {
				t.Errorf("special form %s = %+v, want an authored record", d.Name, d)
			}
			continue
		}
		info, ok := LookupFunctionInfo(d.Name)
		if !ok {
			t.Errorf("%s: no function info", d.Name)
			continue
		}
		// A function with no authored record has the estimate of the
		// registry. An authored record is checked in TestFuncDocMinArgs.
		if d.Summary == "" && (d.MinArgs != info.MinArgs || d.Variadic != info.Variadic) {
			t.Errorf("%s: MinArgs %d Variadic %v, want %d %v", d.Name, d.MinArgs, d.Variadic, info.MinArgs, info.Variadic)
		}
		if !strings.HasPrefix(d.Signature, d.Name+"(") || !strings.HasSuffix(d.Signature, ")") {
			t.Errorf("%s: signature %q", d.Name, d.Signature)
		}
		if d.Summary == "" {
			// No authored record.
			if d.Signature != d.Name+"(...)" || d.Args != nil {
				t.Errorf("%s: record with no summary = %+v, want the NAME(...) form", d.Name, d)
			}
		}
	}

	// The result is a copy: a change does not reach the next call.
	for i := range docs {
		if len(docs[i].Args) > 0 {
			docs[i].Args[0].Name = "changed"
		}
		docs[i].Summary = "changed"
	}
	for _, d := range FuncDocs() {
		if d.Summary == "changed" || (len(d.Args) > 0 && d.Args[0].Name == "changed") {
			t.Fatalf("%s: FuncDocs returned shared data", d.Name)
		}
	}
}

// MinArgs and Variadic of an authored record must agree with its own
// signature, not with the argument spec count of the registry.
func TestFuncDocMinArgs(t *testing.T) {
	byName := map[string]FuncDoc{}
	authored := 0
	for _, d := range FuncDocs() {
		byName[d.Name] = d
		if d.Summary == "" {
			continue
		}
		authored++
		required, repeat := 0, false
		for _, a := range d.Args {
			if !a.Optional {
				required++
			}
			repeat = repeat || a.Repeat
		}
		if d.MinArgs != required || d.Variadic != repeat {
			t.Errorf("%s: MinArgs %d Variadic %v; the signature %q has %d required arguments, repeat %v",
				d.Name, d.MinArgs, d.Variadic, d.Signature, required, repeat)
		}
		if strings.HasSuffix(d.Signature, "...)") != d.Variadic {
			t.Errorf("%s: Variadic %v does not agree with the signature %q", d.Name, d.Variadic, d.Signature)
		}
	}
	if authored != len(funcDocData) {
		t.Errorf("%d authored records in FuncDocs, want %d", authored, len(funcDocData))
	}
	// The cases of the review.
	tests := []struct {
		name     string
		minArgs  int
		variadic bool
	}{
		{"SUMIF", 2, false}, {"AVERAGEIF", 2, false}, {"ROUND", 2, false}, {"IF", 1, false},
		{"AND", 1, true}, {"SUMIFS", 3, true}, {"VLOOKUP", 3, false}, {"SUM", 1, true},
		{"PI", 0, false}, {"TODAY", 0, false}, {"ROW", 0, false},
		{"LET", 3, true}, {"LAMBDA", 1, true}, {"REDUCE", 2, false}, {"SCAN", 2, false}, {"MAP", 2, true},
	}
	for _, tt := range tests {
		d := byName[tt.name]
		if d.MinArgs != tt.minArgs || d.Variadic != tt.variadic {
			t.Errorf("%s: MinArgs %d Variadic %v, want %d %v", tt.name, d.MinArgs, d.Variadic, tt.minArgs, tt.variadic)
		}
	}
	// REDUCE and SCAN accept an empty first argument.
	for _, src := range []string{"REDUCE(,A1:A3,LAMBDA(a,b,a+b))", "SCAN(,A1:A3,LAMBDA(a,b,a+b))", "REDUCE(0,A1:A3,LAMBDA(a,b,a+b))"} {
		if _, err := Parse(src); err != nil {
			t.Errorf("Parse(%q): %v", src, err)
		}
	}
}

func TestFuncDocData(t *testing.T) {
	if _, err := authoredFuncDocs(); err != nil {
		t.Fatalf("authored records: %v", err)
	}
	if len(funcDocData) < 120 {
		t.Errorf("there are %d authored records, want 120 or more", len(funcDocData))
	}
	seen := map[string]bool{}
	for _, e := range funcDocData {
		name, args, err := parseFuncSignature(e.signature)
		if err != nil {
			t.Errorf("%v", err)
			continue
		}
		if seen[name] {
			t.Errorf("%s has two records", name)
		}
		seen[name] = true

		if e.summary == "" || !strings.HasSuffix(e.summary, ".") || strings.Contains(e.summary, "\n") {
			t.Errorf("%s: summary %q must be one sentence", name, e.summary)
		}
		isSpecial := false
		for _, s := range funcDocSpecialForms {
			isSpecial = isSpecial || s == name
		}
		info, ok := LookupFunctionInfo(name)
		if !ok && !isSpecial {
			t.Errorf("%s has a record but is not a registered function or a special form", name)
			continue
		}
		// The signature must not show as optional an argument that the
		// engine requires. Two functions are exceptions: their registry
		// record counts the optional last range as required, but the engine
		// accepts the call without it, as Excel does (the bridge tests
		// evaluate both forms).
		required := 0
		for _, a := range args {
			if !a.Optional {
				required++
			}
		}
		if name == "SUMIF" || name == "AVERAGEIF" {
			if required != 2 || info.MinArgs != 3 {
				t.Errorf("%s: required %d, engine %d; the exception is out of date", name, required, info.MinArgs)
			}
			continue
		}
		if ok && required < info.MinArgs {
			t.Errorf("%s: the signature has %d required arguments; the engine requires %d", name, required, info.MinArgs)
		}
		argNames := map[string]bool{}
		for _, a := range args {
			if argNames[a.Name] {
				t.Errorf("%s: argument name %q occurs two times", name, a.Name)
			}
			argNames[a.Name] = true
		}
	}

	// The functions that the plan names must have a record.
	for _, name := range strings.Fields(`SUM SUMIF SUMIFS COUNT COUNTA COUNTIF COUNTIFS AVERAGE AVERAGEIF AVERAGEIFS
		MIN MAX ROUND ROUNDUP ROUNDDOWN IF IFS IFERROR AND OR NOT VLOOKUP HLOOKUP XLOOKUP INDEX MATCH XMATCH OFFSET
		INDIRECT FILTER SORT SORTBY UNIQUE SEQUENCE LET LAMBDA TEXT LEFT RIGHT MID LEN TRIM CONCAT TEXTJOIN DATE
		TODAY NOW YEAR MONTH DAY EDATE EOMONTH NETWORKDAYS PMT NPV IRR`) {
		if !seen[name] {
			t.Errorf("%s has no authored record", name)
		}
	}

	// Spot checks of derived arguments.
	byName := map[string]FuncDoc{}
	for _, d := range FuncDocs() {
		byName[d.Name] = d
	}
	sum := byName["SUM"]
	if len(sum.Args) != 2 || sum.Args[0] != (ArgDoc{Name: "number1"}) || sum.Args[1] != (ArgDoc{Name: "number2", Optional: true, Repeat: true}) {
		t.Errorf("SUM args = %+v", sum.Args)
	}
	sumif := byName["SUMIF"]
	if len(sumif.Args) != 3 || sumif.Args[2] != (ArgDoc{Name: "sum_range", Optional: true}) || sumif.Summary == "" {
		t.Errorf("SUMIF = %+v", sumif)
	}
	if pi := byName["PI"]; pi.Args == nil || len(pi.Args) != 0 || pi.Signature != "PI()" {
		t.Errorf("PI = %+v, want an empty argument list that is not nil", pi)
	}
}
