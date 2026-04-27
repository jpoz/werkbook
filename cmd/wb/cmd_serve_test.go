package main

import (
	"encoding/json"
	"path/filepath"
	"testing"
)

// dispatchOne exercises the in-process dispatcher (skipping the stdio I/O
// loop) so tests stay fast and deterministic. The serve binary is exercised
// indirectly via the same dispatchServe function.
func dispatchOne(t *testing.T, sess *session, op string, params any) serveResponse {
	t.Helper()
	var raw json.RawMessage
	if params != nil {
		b, err := json.Marshal(params)
		if err != nil {
			t.Fatalf("marshal params: %v", err)
		}
		raw = b
	}
	return dispatchServe(sess, serveRequest{ID: op, Op: op, Params: raw})
}

func TestServeCreateApplyCalcReadSaveCycle(t *testing.T) {
	dir := t.TempDir()
	sess := &session{}

	// 1. Create an in-memory workbook with a header and one row of data.
	createSpec := map[string]any{
		"sheets": []string{"Data"},
		"cells": []any{
			map[string]any{"cell": "A1", "value": "name"},
			map[string]any{"cell": "B1", "value": "amount"},
			map[string]any{"cell": "A2", "value": "alpha"},
			map[string]any{"cell": "B2", "value": 10},
		},
	}
	resp := dispatchOne(t, sess, "create", map[string]any{"spec": createSpec})
	if !resp.OK {
		t.Fatalf("create failed: %+v", resp.Error)
	}

	// 2. Apply a patch that adds a SUM formula.
	resp = dispatchOne(t, sess, "apply", map[string]any{
		"ops": []any{
			map[string]any{"cell": "A3", "value": "total"},
			map[string]any{"cell": "B3", "formula": "SUM(B2:B2)"},
		},
	})
	if !resp.OK {
		t.Fatalf("apply failed: %+v", resp.Error)
	}

	// 3. Recalculate so cached values reflect the new formula.
	resp = dispatchOne(t, sess, "calc", nil)
	if !resp.OK {
		t.Fatalf("calc failed: %+v", resp.Error)
	}

	// 4. Read B3 and check the SUM evaluated to 10.
	resp = dispatchOne(t, sess, "read", map[string]any{"sheet": "Data", "range": "B3:B3"})
	if !resp.OK {
		t.Fatalf("read failed: %+v", resp.Error)
	}
	b, _ := json.Marshal(resp.Data)
	var rd readData
	if err := json.Unmarshal(b, &rd); err != nil {
		t.Fatalf("decode read data: %v", err)
	}
	got := rd.Rows[0].Cells["B3"].Value
	if v, ok := got.(float64); !ok || v != 10 {
		t.Fatalf("expected B3=10, got %v (%T)", got, got)
	}

	// 5. Save to disk and reopen to confirm the file is well-formed.
	out := filepath.Join(dir, "session.xlsx")
	resp = dispatchOne(t, sess, "save", map[string]any{"path": out})
	if !resp.OK {
		t.Fatalf("save failed: %+v", resp.Error)
	}

	sess2 := &session{}
	resp = dispatchOne(t, sess2, "open", map[string]any{"path": out})
	if !resp.OK {
		t.Fatalf("reopen failed: %+v", resp.Error)
	}
	resp = dispatchOne(t, sess2, "info", nil)
	if !resp.OK {
		t.Fatalf("info failed: %+v", resp.Error)
	}
	infoB, _ := json.Marshal(resp.Data)
	var info infoData
	json.Unmarshal(infoB, &info)
	if len(info.Sheets) != 1 || info.Sheets[0].Name != "Data" {
		t.Errorf("unexpected sheets: %+v", info.Sheets)
	}
}

func TestServeRequiresFileLoaded(t *testing.T) {
	sess := &session{}
	for _, op := range []string{"info", "read", "apply", "calc", "save"} {
		resp := dispatchOne(t, sess, op, nil)
		if resp.OK {
			t.Errorf("expected %s to fail without a loaded file", op)
			continue
		}
		if resp.Error == nil || resp.Error.Code != ErrCodeNoFileLoaded {
			t.Errorf("expected NO_FILE_LOADED for %s, got %+v", op, resp.Error)
		}
	}
}

func TestServeUnknownOp(t *testing.T) {
	sess := &session{}
	resp := dispatchOne(t, sess, "frobnicate", nil)
	if resp.OK {
		t.Fatal("expected unknown op to fail")
	}
	if resp.Error.Code != ErrCodeInvalidSpec {
		t.Errorf("expected INVALID_SPEC, got %q", resp.Error.Code)
	}
}

func TestServeApplyPartialFailureFlagsState(t *testing.T) {
	sess := &session{}
	resp := dispatchOne(t, sess, "create", map[string]any{
		"spec": map[string]any{"sheets": []string{"S"}},
	})
	if !resp.OK {
		t.Fatalf("create failed: %+v", resp.Error)
	}
	// Apply two ops where the second targets a bogus reference.
	resp = dispatchOne(t, sess, "apply", map[string]any{
		"ops": []any{
			map[string]any{"cell": "A1", "value": "ok"},
			map[string]any{"cell": "ZZZ1", "value": "boom"},
		},
	})
	if resp.OK {
		t.Fatal("expected apply to report partial failure")
	}
	if resp.Error.Code != ErrCodePartialFailure {
		t.Errorf("expected PARTIAL_FAILURE, got %q", resp.Error.Code)
	}
}

func TestServeQuitDoesNotRequireFile(t *testing.T) {
	sess := &session{}
	resp := dispatchOne(t, sess, "quit", nil)
	if !resp.OK {
		t.Fatalf("quit should always succeed: %+v", resp.Error)
	}
}
