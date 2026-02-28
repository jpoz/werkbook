package fuzz

import (
	"encoding/json"
	"fmt"
	"math/rand"
	"strings"
)

// MutateSpec takes a passing spec and applies targeted mutations to test edge cases.
// Returns a new spec with mutated values (the original spec is not modified).
func MutateSpec(spec *TestSpec) *TestSpec {
	// Deep copy the spec via JSON round-trip.
	data, err := json.Marshal(spec)
	if err != nil {
		return spec
	}
	var mutated TestSpec
	if err := json.Unmarshal(data, &mutated); err != nil {
		return spec
	}

	mutated.Name = fmt.Sprintf("mutated_%s", spec.Name)

	// Apply mutations to cells.
	for si := range mutated.Sheets {
		for ci := range mutated.Sheets[si].Cells {
			cell := &mutated.Sheets[si].Cells[ci]
			if cell.Formula != "" {
				// Don't mutate formulas, only data cells.
				continue
			}
			// Apply mutation with 40% probability per cell.
			if rand.Float64() < 0.4 {
				mutateCell(cell)
			}
		}
	}

	// Clear expected values since mutations change results.
	for i := range mutated.Checks {
		mutated.Checks[i].Expected = ""
	}

	return &mutated
}

// mutateCell applies a random mutation to a data cell based on its type.
func mutateCell(cell *CellSpec) {
	switch cell.Type {
	case "number":
		mutateNumber(cell)
	case "string":
		mutateString(cell)
	case "bool":
		mutateBool(cell)
	default:
		// Unknown type, try numeric mutation.
		mutateNumber(cell)
	}
}

// Edge case numbers for mutation testing.
var edgeNumbers = []float64{
	0, -1, 1, -0.5, 0.5,
	1e15, -1e15, 1e-15, -1e-15,
	0.1, 0.01, 0.001,
	999999999, -999999999,
	2147483647,  // max int32
	-2147483648, // min int32
	0.9999999999,
	3.14159265358979,
}

func mutateNumber(cell *CellSpec) {
	cell.Value = edgeNumbers[rand.Intn(len(edgeNumbers))]
}

// Edge case strings for mutation testing.
var edgeStrings = []string{
	"",           // empty
	" ",          // single space
	"  \t  ",     // whitespace
	"0",          // numeric string
	"123",        // integer string
	"3.14",       // float string
	"TRUE",       // boolean string
	"FALSE",      // boolean string
	"-1",         // negative number string
	"1E5",        // scientific notation string
	"hello world", // space in string
}

func mutateString(cell *CellSpec) {
	cell.Value = edgeStrings[rand.Intn(len(edgeStrings))]
}

func mutateBool(cell *CellSpec) {
	// Flip the boolean.
	if b, ok := cell.Value.(bool); ok {
		cell.Value = !b
	} else {
		cell.Value = true
	}
}

// MutateFormulas wraps formulas in IFERROR() to catch error propagation issues.
func MutateFormulas(spec *TestSpec) *TestSpec {
	data, err := json.Marshal(spec)
	if err != nil {
		return spec
	}
	var mutated TestSpec
	if err := json.Unmarshal(data, &mutated); err != nil {
		return spec
	}

	mutated.Name = fmt.Sprintf("iferror_%s", spec.Name)

	for si := range mutated.Sheets {
		for ci := range mutated.Sheets[si].Cells {
			cell := &mutated.Sheets[si].Cells[ci]
			if cell.Formula != "" && !strings.HasPrefix(strings.ToUpper(cell.Formula), "IFERROR(") {
				cell.Formula = fmt.Sprintf("IFERROR(%s,\"ERR\")", cell.Formula)
			}
		}
	}

	// Clear expected values.
	for i := range mutated.Checks {
		mutated.Checks[i].Expected = ""
	}

	return &mutated
}
