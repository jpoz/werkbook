package main

import (
	"io/fs"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"testing/fstest"
)

func TestHandler(t *testing.T) {
	built := fstest.MapFS{
		"index.html":    {Data: []byte("<!doctype html><title>Werkbook</title>")},
		"app.js":        {Data: []byte("console.log('app')")},
		"style.css":     {Data: []byte("body{}")},
		"werkbook.wasm": {Data: []byte("\x00asm\x01\x00\x00\x00")},
		"wasm_exec.js":  {Data: []byte("// wasm_exec")},
		"js/main.js":    {Data: []byte("import { boot } from './workbook.js';")},
		"css/fill.css":  {Data: []byte("#fill-handle{}")},
	}
	notBuilt := fstest.MapFS{
		"index.html": built["index.html"],
		"app.js":     built["app.js"],
	}

	tests := []struct {
		name       string
		fsys       fs.FS
		method     string
		path       string
		wantStatus int
		wantType   string
		wantBody   string // substring
	}{
		{name: "root serves the index page", fsys: built, path: "/", wantStatus: http.StatusOK, wantType: "text/html; charset=utf-8", wantBody: "Werkbook"},
		{name: "script", fsys: built, path: "/app.js", wantStatus: http.StatusOK, wantType: "text/javascript; charset=utf-8", wantBody: "console.log"},
		{name: "style sheet", fsys: built, path: "/style.css", wantStatus: http.StatusOK, wantType: "text/css; charset=utf-8", wantBody: "body"},
		{name: "ES module in a sub-directory", fsys: built, path: "/js/main.js", wantStatus: http.StatusOK, wantType: "text/javascript; charset=utf-8", wantBody: "import"},
		{name: "style sheet in a sub-directory", fsys: built, path: "/css/fill.css", wantStatus: http.StatusOK, wantType: "text/css; charset=utf-8", wantBody: "fill-handle"},
		{name: "wasm has the wasm content type", fsys: built, path: "/werkbook.wasm", wantStatus: http.StatusOK, wantType: "application/wasm", wantBody: "asm"},
		{name: "wasm support script", fsys: built, path: "/wasm_exec.js", wantStatus: http.StatusOK, wantType: "text/javascript; charset=utf-8", wantBody: "wasm_exec"},
		{name: "HEAD", fsys: built, method: http.MethodHead, path: "/werkbook.wasm", wantStatus: http.StatusOK, wantType: "application/wasm"},
		{name: "wasm that is not built", fsys: notBuilt, path: "/werkbook.wasm", wantStatus: http.StatusNotFound, wantType: "text/plain; charset=utf-8", wantBody: `werkbook.wasm is not built`},
		{name: "not-built message says how to build", fsys: notBuilt, path: "/werkbook.wasm", wantStatus: http.StatusNotFound, wantBody: `make web`},
		{name: "support script that is not built", fsys: notBuilt, path: "/wasm_exec.js", wantStatus: http.StatusNotFound, wantBody: `wasm_exec.js is not built`},
		{name: "index still loads when the wasm is not built", fsys: notBuilt, path: "/", wantStatus: http.StatusOK, wantBody: "Werkbook"},
		{name: "unknown file", fsys: built, path: "/nope.txt", wantStatus: http.StatusNotFound},
		{name: "path that leaves the root", fsys: built, path: "/../main.go", wantStatus: http.StatusNotFound},
		{name: "POST is rejected", fsys: built, method: http.MethodPost, path: "/", wantStatus: http.StatusMethodNotAllowed},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			method := tt.method
			if method == "" {
				method = http.MethodGet
			}
			req := httptest.NewRequest(method, "http://localhost/", nil)
			req.URL.Path = tt.path
			rec := httptest.NewRecorder()
			newHandler(tt.fsys).ServeHTTP(rec, req)

			if rec.Code != tt.wantStatus {
				t.Fatalf("status = %d, want %d (body %q)", rec.Code, tt.wantStatus, rec.Body.String())
			}
			if tt.wantType != "" {
				if got := rec.Header().Get("Content-Type"); got != tt.wantType {
					t.Errorf("Content-Type = %q, want %q", got, tt.wantType)
				}
			}
			if tt.wantBody != "" && !strings.Contains(rec.Body.String(), tt.wantBody) {
				t.Errorf("body = %q, want it to contain %q", rec.Body.String(), tt.wantBody)
			}
		})
	}
}

// The embedded directory must have the index page in all builds, also when
// the WASM module is not built.
func TestEmbeddedStaticHasIndex(t *testing.T) {
	fsys, _, err := staticFS("")
	if err != nil {
		t.Fatalf("staticFS: %v", err)
	}
	if _, err := fs.Stat(fsys, "index.html"); err != nil {
		t.Fatalf("embedded index.html: %v", err)
	}
}

func TestStaticFSDevDir(t *testing.T) {
	if _, _, err := staticFS(t.TempDir()); err != nil {
		t.Errorf("staticFS with a valid directory: %v", err)
	}
	if _, _, err := staticFS("does-not-exist"); err == nil {
		t.Errorf("staticFS with a missing directory gave no error")
	}
	if _, _, err := staticFS("main.go"); err == nil {
		t.Errorf("staticFS with a file gave no error")
	}
}
