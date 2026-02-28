package fuzz

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"time"
)

// TestCase bundles a spec with its oracle ground truth.
type TestCase struct {
	Spec        *TestSpec    `json:"spec"`
	GroundTruth []CellResult `json:"ground_truth"`
	OracleName  string       `json:"oracle_name,omitempty"`
}

// CellResult holds the evaluated value for a single cell.
type CellResult struct {
	Ref   string `json:"ref"`
	Value string `json:"value"`
	Type  string `json:"type"`
}

// CheckResult records the outcome of running the checker on a test case.
type CheckResult struct {
	Passed     bool       `json:"passed"`
	Mismatches []Mismatch `json:"mismatches,omitempty"`
	FixApplied bool       `json:"fix_applied"`
}

// NewTestCaseDir creates a timestamped directory for a test case under outDir.
func NewTestCaseDir(outDir string) (string, error) {
	name := time.Now().Format("20060102_150405")
	dir := filepath.Join(outDir, name)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return "", fmt.Errorf("create test case dir: %w", err)
	}
	return dir, nil
}

// SaveTestCase writes the spec, XLSX, and ground truth to a test case directory.
// oracleName is optional; if provided, it is saved in the test case metadata.
func SaveTestCase(dir string, spec *TestSpec, xlsxPath string, groundTruth []CellResult, oracleName ...string) error {
	// Save spec.
	specPath := filepath.Join(dir, "spec.json")
	if err := SaveSpec(specPath, spec); err != nil {
		return fmt.Errorf("save spec: %w", err)
	}

	// Copy XLSX.
	xlsxDst := filepath.Join(dir, "test.xlsx")
	data, err := os.ReadFile(xlsxPath)
	if err != nil {
		return fmt.Errorf("read xlsx: %w", err)
	}
	if err := os.WriteFile(xlsxDst, data, 0644); err != nil {
		return fmt.Errorf("write xlsx: %w", err)
	}

	// Save ground truth.
	gtPath := filepath.Join(dir, "ground_truth.json")
	gtData, err := json.MarshalIndent(groundTruth, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal ground truth: %w", err)
	}
	if err := os.WriteFile(gtPath, gtData, 0644); err != nil {
		return fmt.Errorf("write ground truth: %w", err)
	}

	// Save oracle name metadata if provided.
	if len(oracleName) > 0 && oracleName[0] != "" {
		meta := map[string]string{"oracle_name": oracleName[0]}
		metaData, err := json.MarshalIndent(meta, "", "  ")
		if err != nil {
			return fmt.Errorf("marshal metadata: %w", err)
		}
		metaPath := filepath.Join(dir, "metadata.json")
		if err := os.WriteFile(metaPath, metaData, 0644); err != nil {
			return fmt.Errorf("write metadata: %w", err)
		}
	}

	return nil
}

// LoadTestCase reads a test case from its directory.
func LoadTestCase(dir string) (*TestCase, error) {
	specPath := filepath.Join(dir, "spec.json")
	spec, err := LoadSpec(specPath)
	if err != nil {
		return nil, fmt.Errorf("load spec: %w", err)
	}

	gtPath := filepath.Join(dir, "ground_truth.json")
	gtData, err := os.ReadFile(gtPath)
	if err != nil {
		return nil, fmt.Errorf("read ground truth: %w", err)
	}
	var groundTruth []CellResult
	if err := json.Unmarshal(gtData, &groundTruth); err != nil {
		return nil, fmt.Errorf("parse ground truth: %w", err)
	}

	tc := &TestCase{
		Spec:        spec,
		GroundTruth: groundTruth,
	}

	// Try to load oracle name from metadata.
	metaPath := filepath.Join(dir, "metadata.json")
	if metaData, err := os.ReadFile(metaPath); err == nil {
		var meta map[string]string
		if err := json.Unmarshal(metaData, &meta); err == nil {
			tc.OracleName = meta["oracle_name"]
		}
	}

	return tc, nil
}

// SaveCheckResult writes the check result to the test case directory.
func SaveCheckResult(dir string, result *CheckResult) error {
	path := filepath.Join(dir, "result.json")
	data, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal result: %w", err)
	}
	return os.WriteFile(path, data, 0644)
}
