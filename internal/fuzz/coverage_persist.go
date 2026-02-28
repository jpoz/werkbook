package fuzz

import (
	"encoding/json"
	"os"
	"time"
)

// coverageSaveData is the on-disk representation of coverage state.
type coverageSaveData struct {
	Tested  map[string]*FunctionStats `json:"tested"`
	Broken  map[string]int            `json:"broken"`
	SavedAt time.Time                 `json:"saved_at"`
}

// SaveCoverage persists coverage data to a JSON file atomically.
func SaveCoverage(path string, fc *FunctionCoverage) error {
	fc.mu.Lock()
	data := coverageSaveData{
		Tested:  fc.Tested,
		Broken:  fc.Broken,
		SavedAt: time.Now(),
	}
	fc.mu.Unlock()

	b, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return err
	}

	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, b, 0644); err != nil {
		return err
	}
	return os.Rename(tmp, path)
}

// LoadCoverage loads coverage data from a JSON file.
// Returns nil, nil if the file does not exist.
func LoadCoverage(path string) (*FunctionCoverage, error) {
	b, err := os.ReadFile(path)
	if os.IsNotExist(err) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}

	var data coverageSaveData
	if err := json.Unmarshal(b, &data); err != nil {
		return nil, err
	}

	fc := NewFunctionCoverage()
	if data.Tested != nil {
		fc.Tested = data.Tested
	}
	if data.Broken != nil {
		fc.Broken = data.Broken
	}
	return fc, nil
}

// MergeCoverage additively merges src into dst.
// Tested counts are summed; Broken values are taken from src.
func MergeCoverage(dst, src *FunctionCoverage) {
	dst.mu.Lock()
	defer dst.mu.Unlock()

	for fn, stats := range src.Tested {
		if existing, ok := dst.Tested[fn]; ok {
			existing.Passed += stats.Passed
			existing.Failed += stats.Failed
		} else {
			dst.Tested[fn] = &FunctionStats{
				Passed: stats.Passed,
				Failed: stats.Failed,
			}
		}
	}

	for fn, count := range src.Broken {
		dst.Broken[fn] = count
	}
}
