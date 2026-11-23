// Command web serves the in-browser XLSX editor. The werkbook library runs in
// the browser as WebAssembly; this server only serves static files. See
// README.md in this directory.
package main

import (
	"context"
	"embed"
	"errors"
	"flag"
	"fmt"
	"io/fs"
	"log"
	"mime"
	"net"
	"net/http"
	"os"
	"os/signal"
	"path"
	"strings"
	"syscall"
	"time"
)

// The directory is embedded, not the build outputs in it. This keeps
// "go build ./..." valid when werkbook.wasm is not built.
//
//go:embed static
var embedded embed.FS

// buildOutputs are the static files that "make web" makes. They are not in
// the repository, so a request for one of them can fail on a new checkout.
var buildOutputs = map[string]bool{
	"werkbook.wasm": true,
	"wasm_exec.js":  true,
}

const notBuiltMessage = "%s is not built.\n\n" +
	"Run \"make web\" in the repository root, then start the server again.\n" +
	"If the server uses -dev, a restart is not necessary.\n"

func main() {
	addr := flag.String("addr", "localhost:8080", "address to listen on")
	dev := flag.String("dev", "", "serve the static files from this `dir` instead of the embedded copy")
	flag.Parse()

	if err := run(*addr, *dev); err != nil {
		log.Fatalf("werkbook-web: %v", err)
	}
}

func run(addr, devDir string) error {
	fsys, source, err := staticFS(devDir)
	if err != nil {
		return err
	}

	ln, err := net.Listen("tcp", addr)
	if err != nil {
		return err
	}
	srv := &http.Server{
		Handler:           newHandler(fsys),
		ReadHeaderTimeout: 10 * time.Second,
	}

	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGTERM)
	defer stop()

	serveErr := make(chan error, 1)
	go func() { serveErr <- srv.Serve(ln) }()

	log.Printf("werkbook-web: serving %s on http://%s", source, ln.Addr())
	if _, err := fs.Stat(fsys, "werkbook.wasm"); err != nil {
		log.Printf("werkbook-web: werkbook.wasm is not built; run \"make web\"")
	}

	select {
	case err := <-serveErr:
		return err
	case <-ctx.Done():
	}
	// A second signal stops the process immediately.
	stop()

	log.Printf("werkbook-web: shutting down")
	shutdownCtx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := srv.Shutdown(shutdownCtx); err != nil {
		return fmt.Errorf("shutdown: %w", err)
	}
	if err := <-serveErr; err != nil && !errors.Is(err, http.ErrServerClosed) {
		return err
	}
	return nil
}

// staticFS returns the file system to serve and a description for the log.
func staticFS(devDir string) (fs.FS, string, error) {
	if devDir == "" {
		sub, err := fs.Sub(embedded, "static")
		if err != nil {
			return nil, "", err
		}
		return sub, "embedded static files", nil
	}
	info, err := os.Stat(devDir)
	if err != nil {
		return nil, "", fmt.Errorf("-dev: %w", err)
	}
	if !info.IsDir() {
		return nil, "", fmt.Errorf("-dev: %s is not a directory", devDir)
	}
	return os.DirFS(devDir), devDir, nil
}

// newHandler serves the static files of fsys.
func newHandler(fsys fs.FS) http.Handler {
	files := http.FileServerFS(fsys)
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet && r.Method != http.MethodHead {
			w.Header().Set("Allow", "GET, HEAD")
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}

		name := strings.TrimPrefix(path.Clean("/"+r.URL.Path), "/")
		if buildOutputs[name] {
			if _, err := fs.Stat(fsys, name); err != nil {
				http.Error(w, fmt.Sprintf(notBuiltMessage, name), http.StatusNotFound)
				return
			}
		}

		// Set the type here so that it does not depend on the MIME table of
		// the host. Browsers refuse WebAssembly.instantiateStreaming for all
		// types but application/wasm.
		if ctype := contentType(name); ctype != "" {
			w.Header().Set("Content-Type", ctype)
		}
		// The files are small and local. Revalidation on each load makes a
		// rebuilt werkbook.wasm or an edited script show immediately.
		w.Header().Set("Cache-Control", "no-cache")
		files.ServeHTTP(w, r)
	})
}

func contentType(name string) string {
	switch path.Ext(name) {
	case ".wasm":
		return "application/wasm"
	case ".js":
		return "text/javascript; charset=utf-8"
	case ".css":
		return "text/css; charset=utf-8"
	case ".html":
		return "text/html; charset=utf-8"
	case "":
		return "" // a directory; the file server serves its index.html
	default:
		return mime.TypeByExtension(path.Ext(name))
	}
}
