package main

import (
	"archive/zip"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"time"
)

// ─── Configuration ──────────────────────────────────────────────────────────

const (
	// SheetJS
	sheetjsTreeURL = "https://api.github.com/repos/SheetJS/test_files/git/trees/master?recursive=1"
	sheetjsRawBase = "https://raw.githubusercontent.com/SheetJS/test_files/master/"

	// GovDocs1
	govdocsZipBase    = "https://downloads.digitalcorpora.org/corpora/files/govdocs1/zipfiles/"
	govdocsSubsetBase = "https://downloads.digitalcorpora.org/corpora/files/govdocs1/threads/"
	govdocsByTypeBase = "https://downloads.digitalcorpora.org/corpora/files/govdocs1/by_type/"

	defaultWorkers = 10
)

// ─── Types ──────────────────────────────────────────────────────────────────

type GitTree struct {
	Tree []TreeEntry `json:"tree"`
}

type TreeEntry struct {
	Path string `json:"path"`
	Type string `json:"type"`
	Size int    `json:"size"`
}

// ─── Main ───────────────────────────────────────────────────────────────────

func main() {
	var (
		xlsxOnly      bool
		workers       int
		sourceSheetJS bool
		sourceGovDocs bool
		govdocsMode   string
		outputDir     string
	)

	flag.BoolVar(&xlsxOnly, "xlsx-only", false, "Only download .xlsx files (skip .xls, .xlsm, .xlsb)")
	flag.IntVar(&workers, "workers", defaultWorkers, "Number of concurrent download workers")
	flag.BoolVar(&sourceSheetJS, "sheetjs", false, "Download from SheetJS/test_files repo")
	flag.BoolVar(&sourceGovDocs, "govdocs", false, "Download from GovDocs1 corpus")
	flag.StringVar(&govdocsMode, "govdocs-mode", "subsets", "GovDocs1 mode: 'subsets' (10 zips, ~10K files), 'full' (all 1000 zips, ~1M files), or 'by-type' (print URLs for pre-sorted archives)")
	flag.StringVar(&outputDir, "output", "xlsx_files", "Output directory")
	flag.Parse()

	// Default: all sources if none specified
	if !sourceSheetJS && !sourceGovDocs {
		sourceSheetJS = true
		sourceGovDocs = true
	}

	extensions := map[string]bool{
		".xlsx": true,
		".xlsm": true,
		".xlsb": true,
		".xls":  true,
	}
	if xlsxOnly {
		extensions = map[string]bool{".xlsx": true}
	}

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		fatal("creating output dir: %v", err)
	}

	var totalSuccess, totalFail int64

	if sourceSheetJS {
		s, f := downloadSheetJS(outputDir, extensions, workers)
		totalSuccess += s
		totalFail += f
	}

	if sourceGovDocs {
		s, f := downloadGovDocs(outputDir, extensions, workers, govdocsMode)
		totalSuccess += s
		totalFail += f
	}

	fmt.Printf("\n══════════════════════════════════════\n")
	fmt.Printf("Total: %d files downloaded, %d failures\n", totalSuccess, totalFail)
	fmt.Printf("Output: %s/\n", outputDir)
}

// ─── SheetJS ────────────────────────────────────────────────────────────────

func downloadSheetJS(outputDir string, extensions map[string]bool, workers int) (int64, int64) {
	fmt.Println("\n── SheetJS/test_files ──────────────────")
	fmt.Println("Fetching file tree from GitHub...")

	client := &http.Client{Timeout: 30 * time.Second}

	req, err := http.NewRequest("GET", sheetjsTreeURL, nil)
	if err != nil {
		fmt.Fprintf(os.Stderr, "  error: %v\n", err)
		return 0, 0
	}

	if token := os.Getenv("GITHUB_TOKEN"); token != "" {
		req.Header.Set("Authorization", "Bearer "+token)
		fmt.Println("  Using GITHUB_TOKEN for authentication")
	} else {
		fmt.Println("  Tip: set GITHUB_TOKEN to avoid rate limits")
	}

	resp, err := client.Do(req)
	if err != nil {
		fmt.Fprintf(os.Stderr, "  error fetching tree: %v\n", err)
		return 0, 0
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		body, _ := io.ReadAll(resp.Body)
		fmt.Fprintf(os.Stderr, "  GitHub API returned %d: %s\n", resp.StatusCode, string(body))
		return 0, 0
	}

	var tree GitTree
	if err := json.NewDecoder(resp.Body).Decode(&tree); err != nil {
		fmt.Fprintf(os.Stderr, "  error decoding: %v\n", err)
		return 0, 0
	}

	var files []TreeEntry
	for _, entry := range tree.Tree {
		if entry.Type != "blob" {
			continue
		}
		ext := strings.ToLower(filepath.Ext(entry.Path))
		if extensions[ext] {
			files = append(files, entry)
		}
	}

	fmt.Printf("  Found %d spreadsheet files\n", len(files))
	if len(files) == 0 {
		return 0, 0
	}

	destDir := filepath.Join(outputDir, "sheetjs")
	os.MkdirAll(destDir, 0755)

	var success, fail int64
	var wg sync.WaitGroup
	ch := make(chan TreeEntry, len(files))

	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			dlClient := &http.Client{Timeout: 60 * time.Second}
			for entry := range ch {
				url := sheetjsRawBase + entry.Path
				safeName := strings.ReplaceAll(entry.Path, "/", "_")
				outPath := filepath.Join(destDir, safeName)

				if err := downloadToFile(dlClient, url, outPath); err != nil {
					atomic.AddInt64(&fail, 1)
					fmt.Printf("  FAIL: %s (%v)\n", entry.Path, err)
				} else {
					s := atomic.AddInt64(&success, 1)
					if s%50 == 0 {
						fmt.Printf("  Progress: %d/%d\n", s, len(files))
					}
				}
			}
		}()
	}

	for _, f := range files {
		ch <- f
	}
	close(ch)
	wg.Wait()

	fmt.Printf("  SheetJS: %d downloaded, %d failed\n", success, fail)
	return success, fail
}

// ─── GovDocs1 ───────────────────────────────────────────────────────────────

func downloadGovDocs(outputDir string, extensions map[string]bool, workers int, mode string) (int64, int64) {
	fmt.Println("\n── GovDocs1 ───────────────────────────")

	destDir := filepath.Join(outputDir, "govdocs1")
	os.MkdirAll(destDir, 0755)

	tmpDir := filepath.Join(outputDir, ".govdocs_tmp")
	os.MkdirAll(tmpDir, 0755)
	defer os.RemoveAll(tmpDir)

	var zipURLs []string

	switch mode {
	case "subsets":
		fmt.Println("  Mode: subsets (10 zips × ~1000 files each)")
		for i := 0; i < 10; i++ {
			zipURLs = append(zipURLs, fmt.Sprintf("%ssubset%d.zip", govdocsSubsetBase, i))
		}
	case "full":
		fmt.Println("  Mode: full (1000 zips × ~1000 files each)")
		fmt.Println("  ⚠  WARNING: This will download ~500GB+ of data!")
		fmt.Println("  Press Ctrl+C within 5 seconds to cancel...")
		time.Sleep(5 * time.Second)
		for i := 0; i < 1000; i++ {
			zipURLs = append(zipURLs, fmt.Sprintf("%s%03d.zip", govdocsZipBase, i))
		}
	case "by-type":
		return downloadGovDocsByType(destDir, extensions)
	default:
		fmt.Fprintf(os.Stderr, "  Unknown govdocs-mode: %s (use 'subsets', 'full', or 'by-type')\n", mode)
		return 0, 0
	}

	var totalSuccess, totalFail int64

	// Process zips with limited concurrency (don't hammer the server)
	zipWorkers := workers
	if zipWorkers > 3 {
		zipWorkers = 3
	}

	var wg sync.WaitGroup
	zipCh := make(chan string, len(zipURLs))

	for i := 0; i < zipWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			dlClient := &http.Client{Timeout: 10 * time.Minute}

			for zipURL := range zipCh {
				zipName := filepath.Base(zipURL)
				zipPath := filepath.Join(tmpDir, zipName)

				fmt.Printf("  Downloading %s...\n", zipName)
				if err := downloadToFile(dlClient, zipURL, zipPath); err != nil {
					fmt.Printf("  FAIL downloading %s: %v\n", zipName, err)
					atomic.AddInt64(&totalFail, 1)
					continue
				}

				extracted, failed := extractFromZip(zipPath, destDir, extensions)
				atomic.AddInt64(&totalSuccess, int64(extracted))
				atomic.AddInt64(&totalFail, int64(failed))

				if extracted > 0 {
					fmt.Printf("  Extracted %d spreadsheets from %s\n", extracted, zipName)
				} else {
					fmt.Printf("  No spreadsheets in %s\n", zipName)
				}

				// Clean up zip to save disk space
				os.Remove(zipPath)
			}
		}()
	}

	for _, url := range zipURLs {
		zipCh <- url
	}
	close(zipCh)
	wg.Wait()

	fmt.Printf("  GovDocs1: %d spreadsheets extracted, %d failures\n", totalSuccess, totalFail)
	return totalSuccess, totalFail
}

func downloadGovDocsByType(destDir string, extensions map[string]bool) (int64, int64) {
	client := &http.Client{Timeout: 30 * time.Second}

	fmt.Println("  Mode: by-type (checking for pre-sorted archives)")
	fmt.Println()

	for ext := range extensions {
		extName := strings.TrimPrefix(ext, ".")

		// Try common archive patterns
		urls := []string{
			fmt.Sprintf("%sfiles.%s.tar", govdocsByTypeBase, extName),
			fmt.Sprintf("%sfiles.%s.zip", govdocsByTypeBase, extName),
		}

		found := false
		for _, url := range urls {
			resp, err := client.Head(url)
			if err != nil {
				continue
			}
			resp.Body.Close()

			if resp.StatusCode == 200 {
				size := resp.ContentLength
				sizeStr := "unknown size"
				if size > 0 {
					sizeStr = formatBytes(size)
				}
				fmt.Printf("  ✓ .%s → %s (%s)\n", extName, url, sizeStr)
				fmt.Printf("    curl -O %s\n", url)
				found = true
				break
			}
		}
		if !found {
			fmt.Printf("  ✗ .%s → no pre-sorted archive found\n", extName)
		}
	}

	fmt.Println()
	fmt.Println("  by-type archives are very large. Download manually with curl,")
	fmt.Println("  or use --govdocs-mode=subsets for automated extraction.")
	return 0, 0
}

// ─── Helpers ────────────────────────────────────────────────────────────────

func downloadToFile(client *http.Client, url, outPath string) error {
	resp, err := client.Get(url)
	if err != nil {
		return fmt.Errorf("GET failed: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return fmt.Errorf("HTTP %d", resp.StatusCode)
	}

	f, err := os.Create(outPath)
	if err != nil {
		return fmt.Errorf("create: %w", err)
	}
	defer f.Close()

	if _, err := io.Copy(f, resp.Body); err != nil {
		return fmt.Errorf("write: %w", err)
	}
	return nil
}

func extractFromZip(zipPath, destDir string, extensions map[string]bool) (extracted, failed int) {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		fmt.Printf("  Error opening zip %s: %v\n", zipPath, err)
		return 0, 1
	}
	defer r.Close()

	for _, f := range r.File {
		if f.FileInfo().IsDir() {
			continue
		}
		ext := strings.ToLower(filepath.Ext(f.Name))
		if !extensions[ext] {
			continue
		}

		zipBase := strings.TrimSuffix(filepath.Base(zipPath), ".zip")
		safeName := fmt.Sprintf("govdocs_%s_%s", zipBase, filepath.Base(f.Name))
		outPath := filepath.Join(destDir, safeName)

		if err := extractZipFile(f, outPath); err != nil {
			fmt.Printf("  Error extracting %s: %v\n", f.Name, err)
			failed++
		} else {
			extracted++
		}
	}
	return
}

func extractZipFile(f *zip.File, outPath string) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, rc)
	return err
}

func formatBytes(b int64) string {
	const unit = 1024
	if b < unit {
		return fmt.Sprintf("%d B", b)
	}
	div, exp := int64(unit), 0
	for n := b / unit; n >= unit; n /= unit {
		div *= unit
		exp++
	}
	return fmt.Sprintf("%.1f %cB", float64(b)/float64(div), "KMGTPE"[exp])
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "FATAL: "+format+"\n", args...)
	os.Exit(1)
}
