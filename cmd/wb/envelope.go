package main

import "encoding/json"

// Response is the JSON envelope wrapping all command output.
type Response struct {
	OK      bool          `json:"ok"`
	Command string        `json:"command"`
	Data    any           `json:"data"`
	Error   *ErrorInfo    `json:"error,omitempty"`
	Meta    *responseMeta `json:"meta,omitempty"`
}

type responseMeta struct {
	SchemaVersion         string   `json:"schema_version"`
	ToolVersion           string   `json:"tool_version"`
	ElapsedMS             int64    `json:"elapsed_ms,omitempty"`
	Mode                  string   `json:"mode,omitempty"`
	Warnings              []string `json:"warnings,omitempty"`
	NextSuggestedCommands []string `json:"next_suggested_commands,omitempty"`
}

// ErrorInfo describes a structured error with remediation hints.
type ErrorInfo struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Hint    string `json:"hint,omitempty"`
}

func successResponse(command string, data any, meta *responseMeta) *Response {
	return &Response{OK: true, Command: command, Data: data, Meta: meta}
}

func errorResponse(command string, ei *ErrorInfo, meta *responseMeta) *Response {
	return &Response{OK: false, Command: command, Data: nil, Error: ei, Meta: meta}
}

func marshalJSON(v any, compact bool) ([]byte, error) {
	if compact {
		return json.Marshal(v)
	}
	return json.MarshalIndent(v, "", "  ")
}
