package formula

import (
	"fmt"
	"sort"
	"strings"
	"sync"
)

// ArgDoc describes one argument of a function signature.
type ArgDoc struct {
	Name     string
	Optional bool // written in brackets
	Repeat   bool // can occur again (the group before "...")
}

// FuncDoc is the documentation record of one function, for editors that show
// a function list and argument hints.
type FuncDoc struct {
	Name string // upper case
	// Signature is the call form, for example
	// "SUMIF(range, criteria, [sum_range])". A function with no authored
	// record has "NAME(...)".
	Signature string
	// Summary is one sentence, or "" when there is no authored record.
	Summary string
	// Args is derived from Signature; nil when there is no authored record.
	Args []ArgDoc
	// MinArgs is the smallest number of arguments of a call and Variadic
	// reports whether the last arguments can occur again.
	//
	// For a function with an authored record, both come from the signature:
	// MinArgs is the number of arguments that are not optional, and Variadic
	// is true when an argument repeats. For a function with no record they
	// are the values of LookupFunctionInfo, which are an estimate: the
	// registry counts argument specs, not the arguments that a call needs.
	MinArgs  int
	Variadic bool
}

// funcDocEntry is one authored record. The function name is the text of the
// signature before the parenthesis.
type funcDocEntry struct {
	signature string
	summary   string
}

var (
	funcDocOnce    sync.Once
	funcDocByName  map[string]FuncDoc
	funcDocInitErr error
)

func authoredFuncDocs() (map[string]FuncDoc, error) {
	funcDocOnce.Do(func() {
		funcDocByName = make(map[string]FuncDoc, len(funcDocData))
		for _, e := range funcDocData {
			name, args, err := parseFuncSignature(e.signature)
			if err != nil {
				funcDocInitErr = err
				return
			}
			if _, dup := funcDocByName[name]; dup {
				funcDocInitErr = fmt.Errorf("function documentation: %s has two records", name)
				return
			}
			doc := FuncDoc{Name: name, Signature: e.signature, Summary: e.summary, Args: args}
			for _, a := range args {
				if !a.Optional {
					doc.MinArgs++
				}
				doc.Variadic = doc.Variadic || a.Repeat
			}
			funcDocByName[name] = doc
		}
	})
	return funcDocByName, funcDocInitErr
}

// funcDocSpecialForms are the names that a user calls as functions but that
// the parser handles itself (LET, LAMBDA and the LAMBDA helpers). They are
// not in the function registry, so they have no FunctionInfo.
var funcDocSpecialForms = []string{"BYCOL", "BYROW", "LAMBDA", "LET", "MAKEARRAY", "MAP", "REDUCE", "SCAN"}

// FuncDocs returns one record for each name that a formula can call: each
// registered function and each special form (see funcDocSpecialForms),
// sorted by name. The result is a new slice; the caller can change it.
func FuncDocs() []FuncDoc {
	authored, _ := authoredFuncDocs()
	infos := RegisteredFunctionInfos()
	out := make([]FuncDoc, 0, len(infos)+len(funcDocSpecialForms))
	seen := make(map[string]bool, len(infos))
	for _, info := range infos {
		doc, ok := authored[info.Name]
		if ok {
			doc.Args = cloneArgDocs(doc.Args)
		} else {
			doc = FuncDoc{Name: info.Name, Signature: info.Name + "(...)", MinArgs: info.MinArgs, Variadic: info.Variadic}
		}
		out = append(out, doc)
		seen[info.Name] = true
	}
	for _, name := range funcDocSpecialForms {
		doc, ok := authored[name]
		if !ok || seen[name] {
			continue
		}
		doc.Args = cloneArgDocs(doc.Args)
		out = append(out, doc)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Name < out[j].Name })
	return out
}

func cloneArgDocs(args []ArgDoc) []ArgDoc {
	out := make([]ArgDoc, len(args))
	copy(out, args)
	return out
}

// parseFuncSignature reads a signature of the form
//
//	NAME(arg, [optional], repeat1, [repeat2], ...)
//
// A bracket group can hold more than one argument ("[range2, criteria2]").
// The "..." mark makes the group before it a repeat group: its arguments can
// occur again.
func parseFuncSignature(sig string) (name string, args []ArgDoc, err error) {
	open := strings.IndexByte(sig, '(')
	if open <= 0 || !strings.HasSuffix(sig, ")") {
		return "", nil, fmt.Errorf("function signature %q: want NAME(...)", sig)
	}
	name = sig[:open]
	if name != strings.ToUpper(name) || strings.ContainsAny(name, " \t") {
		return "", nil, fmt.Errorf("function signature %q: the name must be upper case", sig)
	}
	body := strings.TrimSpace(sig[open+1 : len(sig)-1])
	if body == "" {
		return name, []ArgDoc{}, nil
	}

	args = []ArgDoc{}
	groupOf := []int{} // group number of each argument
	group := 0
	inBracket := false
	sawRepeat := false
	for _, raw := range strings.Split(body, ",") {
		item := strings.TrimSpace(raw)
		if sawRepeat {
			return "", nil, fmt.Errorf("function signature %q: text after \"...\"", sig)
		}
		if item == "..." {
			if inBracket || len(args) == 0 {
				return "", nil, fmt.Errorf("function signature %q: \"...\" in a wrong position", sig)
			}
			last := groupOf[len(groupOf)-1]
			for i := range args {
				if groupOf[i] == last {
					args[i].Repeat = true
				}
			}
			sawRepeat = true
			continue
		}
		optional := inBracket
		if strings.HasPrefix(item, "[") {
			if inBracket {
				return "", nil, fmt.Errorf("function signature %q: nested brackets", sig)
			}
			inBracket, optional = true, true
			group++
			item = item[1:]
		} else if !inBracket {
			group++
		}
		if strings.HasSuffix(item, "]") {
			if !inBracket {
				return "", nil, fmt.Errorf("function signature %q: \"]\" with no \"[\"", sig)
			}
			inBracket = false
			item = item[:len(item)-1]
		}
		item = strings.TrimSpace(item)
		if !isArgName(item) {
			return "", nil, fmt.Errorf("function signature %q: bad argument name %q", sig, item)
		}
		args = append(args, ArgDoc{Name: item, Optional: optional})
		groupOf = append(groupOf, group)
	}
	if inBracket {
		return "", nil, fmt.Errorf("function signature %q: \"[\" with no \"]\"", sig)
	}
	return name, args, nil
}

func isArgName(s string) bool {
	if s == "" || !isAlpha(s[0]) {
		return false
	}
	for i := 0; i < len(s); i++ {
		if c := s[i]; !isAlpha(c) && !isDigit(c) && c != '_' {
			return false
		}
	}
	return true
}
