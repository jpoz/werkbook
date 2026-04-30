package main

import (
	"bufio"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

// session holds the in-memory state of a wb serve loop. One workbook at a
// time; subsequent open/create calls replace whatever was loaded.
type session struct {
	file *werkbook.File
	path string // last open() path or last save() path; "" if neither
}

// serveRequest is one line of NDJSON on stdin. ID is echoed back so callers
// can correlate; if absent the response uses an empty ID.
type serveRequest struct {
	ID     string          `json:"id,omitempty"`
	Op     string          `json:"op"`
	Params json.RawMessage `json:"params,omitempty"`
}

// serveResponse is one line of NDJSON on stdout.
type serveResponse struct {
	ID    string     `json:"id,omitempty"`
	Op    string     `json:"op,omitempty"`
	OK    bool       `json:"ok"`
	Data  any        `json:"data,omitempty"`
	Error *ErrorInfo `json:"error,omitempty"`
}

// serveOp describes a single supported operation for capabilities.
type serveOp struct {
	Name        string `json:"name"`
	Description string `json:"description"`
	ParamsKind  string `json:"params_kind,omitempty"` // schema name from inputSchemas if applicable
}

func cmdServe(args []string, globals globalFlags) int {
	cmd := "serve"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if len(args) > 0 {
		writeError(cmd, errUsage("serve does not accept positional arguments"), globals)
		return ExitUsage
	}

	in := bufio.NewReader(os.Stdin)
	out := bufio.NewWriter(os.Stdout)
	defer out.Flush()

	enc := json.NewEncoder(out)
	enc.SetEscapeHTML(false)

	sess := &session{}

	for {
		line, err := in.ReadBytes('\n')
		if err != nil {
			if errors.Is(err, io.EOF) {
				if len(strings.TrimSpace(string(line))) == 0 {
					return ExitSuccess
				}
				// trailing line without newline; fall through to dispatch
			} else {
				return ExitInternal
			}
		}

		trimmed := strings.TrimSpace(string(line))
		if trimmed == "" {
			if errors.Is(err, io.EOF) {
				return ExitSuccess
			}
			continue
		}

		var req serveRequest
		if jerr := json.Unmarshal([]byte(trimmed), &req); jerr != nil {
			writeServeResponse(enc, out, serveResponse{
				OK: false,
				Error: &ErrorInfo{
					Code:    ErrCodeInvalidSpec,
					Message: "request is not valid JSON: " + jerr.Error(),
					Hint:    "Send one JSON object per line: {\"id\":\"1\",\"op\":\"capabilities\"}",
				},
			})
			if errors.Is(err, io.EOF) {
				return ExitSuccess
			}
			continue
		}

		resp := dispatchServe(sess, req)
		writeServeResponse(enc, out, resp)

		if req.Op == "quit" && resp.OK {
			return ExitSuccess
		}
		if errors.Is(err, io.EOF) {
			return ExitSuccess
		}
	}
}

func writeServeResponse(enc *json.Encoder, w *bufio.Writer, resp serveResponse) {
	_ = enc.Encode(resp)
	_ = w.Flush()
}

// dispatchServe routes one request to its handler and packages the result
// into a serveResponse. Handlers operate on *session and return (data, *ErrorInfo).
func dispatchServe(sess *session, req serveRequest) serveResponse {
	resp := serveResponse{ID: req.ID, Op: req.Op}
	var data any
	var errInfo *ErrorInfo

	switch req.Op {
	case "capabilities":
		data, errInfo = serveCapabilities(sess, req.Params)
	case "open":
		data, errInfo = serveOpen(sess, req.Params)
	case "create":
		data, errInfo = serveCreate(sess, req.Params)
	case "info":
		data, errInfo = serveInfo(sess, req.Params)
	case "read":
		data, errInfo = serveRead(sess, req.Params)
	case "apply":
		data, errInfo = serveApply(sess, req.Params)
	case "calc":
		data, errInfo = serveCalc(sess, req.Params)
	case "save":
		data, errInfo = serveSave(sess, req.Params)
	case "close":
		data, errInfo = serveClose(sess, req.Params)
	case "quit":
		data = map[string]any{"goodbye": true}
	case "":
		errInfo = &ErrorInfo{Code: ErrCodeInvalidSpec, Message: "missing 'op' field"}
	default:
		errInfo = &ErrorInfo{
			Code:    ErrCodeInvalidSpec,
			Message: fmt.Sprintf("unknown op %q", req.Op),
			Hint:    "Send {\"op\":\"capabilities\"} for the list of supported ops.",
		}
	}

	if errInfo != nil {
		resp.OK = false
		resp.Error = errInfo
		return resp
	}
	resp.OK = true
	resp.Data = data
	return resp
}
