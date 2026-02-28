package fuzz

import (
	"encoding/json"
	"os"
	"path/filepath"
	"sort"
	"time"
)

// FixRecord stores information about a successful fix for future reference.
type FixRecord struct {
	Function     string    `json:"function"`
	FailureType  string    `json:"failure_type"`
	ErrorPattern string    `json:"error_pattern"`
	FixSummary   string    `json:"fix_summary"`
	Timestamp    time.Time `json:"timestamp"`
}

// SaveFixRecord persists a fix record to the history directory.
func SaveFixRecord(dir string, record FixRecord) error {
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}

	name := record.Function + "_" + record.Timestamp.Format("20060102_150405") + ".json"
	path := filepath.Join(dir, name)

	data, err := json.MarshalIndent(record, "", "  ")
	if err != nil {
		return err
	}

	return os.WriteFile(path, data, 0644)
}

// LoadFixHistory loads all fix records from the history directory.
// Returns an empty slice if the directory doesn't exist.
func LoadFixHistory(dir string) []FixRecord {
	matches, err := filepath.Glob(filepath.Join(dir, "*.json"))
	if err != nil || len(matches) == 0 {
		return nil
	}

	var records []FixRecord
	for _, path := range matches {
		data, err := os.ReadFile(path)
		if err != nil {
			continue
		}
		var record FixRecord
		if err := json.Unmarshal(data, &record); err != nil {
			continue
		}
		records = append(records, record)
	}

	// Sort by timestamp, most recent first.
	sort.Slice(records, func(i, j int) bool {
		return records[i].Timestamp.After(records[j].Timestamp)
	})

	return records
}
