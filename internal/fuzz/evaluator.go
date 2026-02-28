package fuzz

import "fmt"

// Evaluator is the interface for formula evaluation oracles.
// Both LibreOffice and MS Graph (Excel Online) implement it.
type Evaluator interface {
	// Name returns a human-readable name for this evaluator (e.g. "libreoffice", "excel").
	Name() string
	// Eval evaluates the XLSX at xlsxPath and returns results for the given checks.
	Eval(xlsxPath string, checks []CheckSpec) ([]CellResult, error)
	// ExcludedFunctions returns functions that should not appear in specs for this evaluator.
	ExcludedFunctions() map[string]bool
}

// NewEvaluator creates an Evaluator by name.
// Supported names: "libreoffice", "excel", "local-excel".
func NewEvaluator(name string) (Evaluator, error) {
	switch name {
	case "libreoffice":
		soffice, err := FindLibreOffice()
		if err != nil {
			return nil, err
		}
		return &LibreOfficeEvaluator{SOfficePath: soffice}, nil
	case "excel":
		return NewMSGraphEvaluator()
	case "local-excel":
		if err := FindLocalExcel(); err != nil {
			return nil, err
		}
		return &LocalExcelEvaluator{}, nil
	default:
		return nil, fmt.Errorf("unknown evaluator: %q (supported: libreoffice, excel, local-excel)", name)
	}
}
