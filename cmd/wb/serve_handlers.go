package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"

	werkbook "github.com/jpoz/werkbook"
)

// requireFile returns an error envelope when the session has no workbook
// loaded. Most ops need one; capabilities/open/create/quit do not.
func requireFile(s *session) *ErrorInfo {
	if s.file == nil {
		return &ErrorInfo{
			Code:    ErrCodeNoFileLoaded,
			Message: "no workbook is loaded in this session",
			Hint:    "Send {\"op\":\"open\",\"params\":{\"path\":\"...\"}} or {\"op\":\"create\",\"params\":{\"spec\":{...}}} first.",
		}
	}
	return nil
}

// serveCapabilities returns the wb top-level toolSpec plus the list of
// session-only ops, so an agent fetches everything in one call.
func serveCapabilities(_ *session, _ json.RawMessage) (any, *ErrorInfo) {
	return map[string]any{
		"tool": wbToolSpec(),
		"session_ops": []serveOp{
			{Name: "capabilities", Description: "Return this object."},
			{Name: "open", Description: "Open an existing .xlsx as the session workbook.", ParamsKind: "open_params"},
			{Name: "create", Description: "Create a new workbook in memory from a create_spec.", ParamsKind: "create_spec"},
			{Name: "info", Description: "Return sheet metadata for the session workbook."},
			{Name: "read", Description: "Read cells from the session workbook.", ParamsKind: "read_params"},
			{Name: "apply", Description: "Apply a patch_array to the session workbook.", ParamsKind: "apply_params"},
			{Name: "calc", Description: "Recalculate every formula in the session workbook."},
			{Name: "save", Description: "Save the session workbook to disk.", ParamsKind: "save_params"},
			{Name: "close", Description: "Discard the in-memory workbook."},
			{Name: "quit", Description: "Terminate the serve loop and exit cleanly."},
		},
	}, nil
}

type openParams struct {
	Path string `json:"path"`
}

func serveOpen(s *session, raw json.RawMessage) (any, *ErrorInfo) {
	var p openParams
	if err := strictUnmarshal(raw, &p); err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
	}
	if p.Path == "" {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: "open requires 'path'"}
	}
	f, err := werkbook.Open(p.Path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, &ErrorInfo{Code: ErrCodeFileNotFound, Message: err.Error()}
		}
		if errors.Is(err, werkbook.ErrEncryptedFile) {
			return nil, &ErrorInfo{Code: ErrCodeEncryptedFile, Message: err.Error()}
		}
		return nil, &ErrorInfo{Code: ErrCodeFileOpenFailed, Message: err.Error()}
	}
	s.file = f
	s.path = p.Path
	return map[string]any{
		"path":   p.Path,
		"sheets": f.SheetNames(),
	}, nil
}

type createParams struct {
	Spec json.RawMessage `json:"spec"`
}

func serveCreate(s *session, raw json.RawMessage) (any, *ErrorInfo) {
	var p createParams
	if err := strictUnmarshal(raw, &p); err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
	}
	if len(p.Spec) == 0 {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: "create requires 'spec'"}
	}

	var spec createSpec
	if err := strictUnmarshal(p.Spec, &spec); err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
	}

	var f *werkbook.File
	if len(spec.Sheets) > 0 {
		f = werkbook.New(werkbook.FirstSheet(spec.Sheets[0]))
		for _, name := range spec.Sheets[1:] {
			f.NewSheet(name)
		}
	} else {
		f = werkbook.New()
	}

	defaultSheet := f.SheetNames()[0]
	rowOps, err := rowsToPatchOps(spec.Rows, defaultSheet)
	if err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
	}
	allOps := append(spec.Cells, rowOps...)
	results, applied := applyPatches(f, allOps, defaultSheet)
	failed := len(allOps) - applied

	if failed > 0 {
		return map[string]any{
			"applied":    applied,
			"failed":     failed,
			"operations": results,
		}, &ErrorInfo{
			Code:    ErrCodePartialFailure,
			Message: fmt.Sprintf("%d of %d operations failed", failed, len(allOps)),
			Hint:    "Workbook was not loaded into the session. Inspect 'data.operations' for errors.",
		}
	}

	s.file = f
	s.path = ""
	return map[string]any{
		"sheets":  f.SheetNames(),
		"applied": applied,
	}, nil
}

func serveInfo(s *session, _ json.RawMessage) (any, *ErrorInfo) {
	if errInfo := requireFile(s); errInfo != nil {
		return nil, errInfo
	}
	data := infoData{File: s.path}
	for _, name := range s.file.SheetNames() {
		sh := s.file.Sheet(name)
		if sh == nil {
			continue
		}
		data.Sheets = append(data.Sheets, buildSheetInfo(sh))
	}
	return data, nil
}

type readParams struct {
	Sheet           string `json:"sheet,omitempty"`
	Range           string `json:"range,omitempty"`
	IncludeFormulas bool   `json:"include_formulas,omitempty"`
	ShowFormulas    bool   `json:"show_formulas,omitempty"`
	IncludeStyles   bool   `json:"include_styles,omitempty"`
	Headers         bool   `json:"headers,omitempty"`
	Limit           int    `json:"limit,omitempty"`
}

func serveRead(s *session, raw json.RawMessage) (any, *ErrorInfo) {
	if errInfo := requireFile(s); errInfo != nil {
		return nil, errInfo
	}
	var p readParams
	if len(raw) > 0 {
		if err := strictUnmarshal(raw, &p); err != nil {
			return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
		}
	}

	sheetName := p.Sheet
	if sheetName == "" {
		names := s.file.SheetNames()
		if len(names) == 0 {
			return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: "workbook has no sheets"}
		}
		sheetName = names[0]
	}
	sh := s.file.Sheet(sheetName)
	if sh == nil {
		return nil, &ErrorInfo{Code: ErrCodeSheetNotFound, Message: fmt.Sprintf("sheet %q not found", sheetName)}
	}

	opts := readOpts{
		rangeFlag:       p.Range,
		headersFlag:     p.Headers,
		includeFormulas: p.IncludeFormulas,
		showFormulas:    p.ShowFormulas,
		includeStyles:   p.IncludeStyles,
		limitFlag:       p.Limit,
	}
	data, exitCode, errInfo := buildReadData(sh, s.path, sheetName, opts)
	if errInfo != nil {
		_ = exitCode
		return nil, errInfo
	}
	return data, nil
}

type applyParams struct {
	Ops    json.RawMessage `json:"ops"`
	Sheet  string          `json:"sheet,omitempty"`
	Atomic *bool           `json:"atomic,omitempty"`
}

func serveApply(s *session, raw json.RawMessage) (any, *ErrorInfo) {
	if errInfo := requireFile(s); errInfo != nil {
		return nil, errInfo
	}
	var p applyParams
	if err := strictUnmarshal(raw, &p); err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
	}
	if len(p.Ops) == 0 {
		return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: "apply requires 'ops'"}
	}
	ops, err := parsePatchOps(p.Ops)
	if err != nil {
		return nil, &ErrorInfo{Code: ErrCodeInvalidPatch, Message: err.Error()}
	}

	atomic := true
	if p.Atomic != nil {
		atomic = *p.Atomic
	}

	defaultSheet := p.Sheet
	if defaultSheet == "" {
		names := s.file.SheetNames()
		if len(names) > 0 {
			defaultSheet = names[0]
		}
	}

	// Atomic mode: apply against a clone in memory; commit only if all succeed.
	// Simpler v1 implementation: apply directly. If atomic and any op fails,
	// the session is left in a partially-mutated state — flag this in the
	// response and recommend close+reopen on failure.
	results, applied := applyPatches(s.file, ops, defaultSheet)
	failed := len(ops) - applied

	data := map[string]any{
		"applied":    applied,
		"failed":     failed,
		"operations": results,
		"atomic":     atomic,
	}
	if failed > 0 {
		hint := "Inspect 'data.operations' for per-operation errors."
		if atomic {
			hint += " The session workbook is now in a partially-mutated state; consider 'close' + 'open' to start fresh."
		}
		return data, &ErrorInfo{
			Code:    ErrCodePartialFailure,
			Message: fmt.Sprintf("%d of %d operations failed", failed, len(ops)),
			Hint:    hint,
		}
	}
	return data, nil
}

func serveCalc(s *session, _ json.RawMessage) (any, *ErrorInfo) {
	if errInfo := requireFile(s); errInfo != nil {
		return nil, errInfo
	}
	s.file.Recalculate()
	return map[string]any{"recalculated": true}, nil
}

type saveParams struct {
	Path string `json:"path,omitempty"`
}

func serveSave(s *session, raw json.RawMessage) (any, *ErrorInfo) {
	if errInfo := requireFile(s); errInfo != nil {
		return nil, errInfo
	}
	var p saveParams
	if len(raw) > 0 {
		if err := strictUnmarshal(raw, &p); err != nil {
			return nil, &ErrorInfo{Code: ErrCodeInvalidSpec, Message: err.Error()}
		}
	}
	dest := p.Path
	if dest == "" {
		dest = s.path
	}
	if dest == "" {
		return nil, &ErrorInfo{
			Code:    ErrCodeInvalidSpec,
			Message: "save requires 'path' (no source path on this session)",
			Hint:    "Workbooks created via 'create' have no source path; pass {\"path\":\"...\"}.",
		}
	}
	if err := s.file.SaveAs(dest); err != nil {
		return nil, &ErrorInfo{Code: ErrCodeFileSaveFailed, Message: err.Error()}
	}
	s.path = dest
	return map[string]any{"path": dest, "saved": true}, nil
}

func serveClose(s *session, _ json.RawMessage) (any, *ErrorInfo) {
	had := s.file != nil
	s.file = nil
	s.path = ""
	return map[string]any{"closed": had}, nil
}
