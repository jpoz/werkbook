package main

import (
	"flag"
	"fmt"
	"io"
	"net/http"
	"os"
	"regexp"
	"sort"
	"strings"
)

const indexURL = "https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188"

func main() {
	listFlag := flag.Bool("list", false, "list all known Excel functions from the index page")
	urlFlag := flag.Bool("url", false, "print just the URL instead of fetching content")
	rawFlag := flag.Bool("raw", false, "print raw extracted text without section formatting")
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, "Usage: exceldoc [flags] [FUNCTION_NAME]\n\n")
		fmt.Fprintf(os.Stderr, "Fetch Excel function documentation from Microsoft.\n\n")
		fmt.Fprintf(os.Stderr, "Examples:\n")
		fmt.Fprintf(os.Stderr, "  exceldoc SUM          # Show SUM documentation\n")
		fmt.Fprintf(os.Stderr, "  exceldoc --list        # List all functions\n")
		fmt.Fprintf(os.Stderr, "  exceldoc --url VLOOKUP # Print doc URL for VLOOKUP\n\n")
		fmt.Fprintf(os.Stderr, "Flags:\n")
		flag.PrintDefaults()
	}
	flag.Parse()

	if *listFlag {
		if err := listFunctions(); err != nil {
			fmt.Fprintf(os.Stderr, "Error: %v\n", err)
			os.Exit(1)
		}
		return
	}

	if flag.NArg() < 1 {
		flag.Usage()
		os.Exit(1)
	}

	funcName := strings.ToUpper(flag.Arg(0))

	if err := fetchDoc(funcName, *urlFlag, *rawFlag); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

func fetchPage(url string) (string, error) {
	resp, err := http.Get(url)
	if err != nil {
		return "", fmt.Errorf("fetch %s: %w", url, err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("fetch %s: status %d", url, resp.StatusCode)
	}

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("read %s: %w", url, err)
	}

	return string(body), nil
}

// parseFunctionLinks extracts function name -> doc URL mappings from the index page HTML.
func parseFunctionLinks(html string) map[string]string {
	result := make(map[string]string)

	// Match links that contain function names. Microsoft's page uses patterns like:
	// <a href="https://support.microsoft.com/...">FUNCTION_NAME function</a>
	// We look for links containing "support.microsoft.com" with function-like text.
	linkRe := regexp.MustCompile(`<a[^>]+href="(https://support\.microsoft\.com/[^"]*)"[^>]*>([^<]+)</a>`)
	matches := linkRe.FindAllStringSubmatch(html, -1)

	for _, m := range matches {
		url := m[1]
		text := strings.TrimSpace(m[2])

		// Extract function name from link text like "SUM function" or "VLOOKUP function"
		text = strings.TrimSuffix(text, " function")
		text = strings.TrimSuffix(text, " Function")

		// Function names are all uppercase (possibly with dots like CEILING.MATH)
		name := strings.TrimSpace(text)
		if name == "" {
			continue
		}

		// Only include entries that look like function names
		if isFunctionName(name) {
			result[strings.ToUpper(name)] = url
		}
	}

	return result
}

// isFunctionName checks if a string looks like an Excel function name.
func isFunctionName(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_') {
			return false
		}
	}
	return true
}

func listFunctions() error {
	fmt.Fprintf(os.Stderr, "Fetching function index...\n")
	html, err := fetchPage(indexURL)
	if err != nil {
		return err
	}

	funcs := parseFunctionLinks(html)
	if len(funcs) == 0 {
		return fmt.Errorf("no functions found on index page")
	}

	// Sort and print.
	names := make([]string, 0, len(funcs))
	for name := range funcs {
		names = append(names, name)
	}
	sort.Strings(names)

	for _, name := range names {
		fmt.Println(name)
	}

	fmt.Fprintf(os.Stderr, "\n%d functions found\n", len(names))
	return nil
}

func fetchDoc(funcName string, urlOnly, raw bool) error {
	fmt.Fprintf(os.Stderr, "Fetching function index...\n")
	html, err := fetchPage(indexURL)
	if err != nil {
		return err
	}

	funcs := parseFunctionLinks(html)
	docURL, ok := funcs[funcName]
	if !ok {
		// Try case-insensitive search.
		for name, url := range funcs {
			if strings.EqualFold(name, funcName) {
				docURL = url
				ok = true
				break
			}
		}
	}

	if !ok {
		return fmt.Errorf("function %q not found in Excel function index", funcName)
	}

	if urlOnly {
		fmt.Println(docURL)
		return nil
	}

	fmt.Fprintf(os.Stderr, "Fetching documentation for %s...\n", funcName)
	docHTML, err := fetchPage(docURL)
	if err != nil {
		return err
	}

	text := extractDocContent(docHTML)
	if raw {
		fmt.Print(text)
	} else {
		printFormatted(funcName, docURL, text)
	}

	return nil
}

// extractDocContent strips HTML and extracts readable text from a doc page.
func extractDocContent(html string) string {
	// Remove script and style tags with their content.
	scriptRe := regexp.MustCompile(`(?is)<(script|style)[^>]*>.*?</\1>`)
	html = scriptRe.ReplaceAllString(html, "")

	// Remove HTML comments.
	commentRe := regexp.MustCompile(`(?s)<!--.*?-->`)
	html = commentRe.ReplaceAllString(html, "")

	// Convert headings to formatted text.
	headingRe := regexp.MustCompile(`(?i)<h[1-6][^>]*>(.*?)</h[1-6]>`)
	html = headingRe.ReplaceAllStringFunc(html, func(match string) string {
		inner := headingRe.FindStringSubmatch(match)
		if len(inner) > 1 {
			text := stripTags(inner[1])
			return "\n## " + text + "\n"
		}
		return match
	})

	// Convert paragraphs to double newlines.
	pRe := regexp.MustCompile(`(?i)</?p[^>]*>`)
	html = pRe.ReplaceAllString(html, "\n")

	// Convert <br> to newlines.
	brRe := regexp.MustCompile(`(?i)<br\s*/?>`)
	html = brRe.ReplaceAllString(html, "\n")

	// Convert list items.
	liRe := regexp.MustCompile(`(?i)<li[^>]*>`)
	html = liRe.ReplaceAllString(html, "\n- ")

	// Strip all remaining HTML tags.
	html = stripTags(html)

	// Decode common HTML entities.
	html = strings.ReplaceAll(html, "&amp;", "&")
	html = strings.ReplaceAll(html, "&lt;", "<")
	html = strings.ReplaceAll(html, "&gt;", ">")
	html = strings.ReplaceAll(html, "&quot;", "\"")
	html = strings.ReplaceAll(html, "&#39;", "'")
	html = strings.ReplaceAll(html, "&nbsp;", " ")
	html = strings.ReplaceAll(html, "&#160;", " ")

	// Collapse multiple blank lines.
	multiNewline := regexp.MustCompile(`\n{3,}`)
	html = multiNewline.ReplaceAllString(html, "\n\n")

	// Trim leading/trailing whitespace per line.
	lines := strings.Split(html, "\n")
	for i, line := range lines {
		lines[i] = strings.TrimSpace(line)
	}

	return strings.TrimSpace(strings.Join(lines, "\n"))
}

// stripTags removes all HTML tags from a string.
func stripTags(s string) string {
	tagRe := regexp.MustCompile(`<[^>]*>`)
	return tagRe.ReplaceAllString(s, "")
}

// printFormatted prints the documentation with nice formatting.
func printFormatted(funcName, url, text string) {
	fmt.Printf("=== %s ===\n", funcName)
	fmt.Printf("URL: %s\n", url)
	fmt.Println(strings.Repeat("-", 60))

	// Try to extract key sections.
	sections := extractSections(text)

	if desc, ok := sections["Description"]; ok {
		fmt.Printf("\nDescription:\n%s\n", desc)
	}
	if syntax, ok := sections["Syntax"]; ok {
		fmt.Printf("\nSyntax:\n%s\n", syntax)
	}
	if remarks, ok := sections["Remarks"]; ok {
		fmt.Printf("\nRemarks:\n%s\n", remarks)
	}
	if example, ok := sections["Example"]; ok {
		fmt.Printf("\nExample:\n%s\n", example)
	}

	// If no sections extracted, print raw text.
	if len(sections) == 0 {
		fmt.Println()
		fmt.Println(text)
	}
}

// extractSections attempts to extract named sections from the doc text.
func extractSections(text string) map[string]string {
	sections := make(map[string]string)
	sectionNames := []string{"Description", "Syntax", "Remarks", "Example", "Examples"}

	lines := strings.Split(text, "\n")
	var currentSection string
	var currentContent []string

	for _, line := range lines {
		// Check if this line is a section header.
		found := false
		for _, name := range sectionNames {
			trimmed := strings.TrimPrefix(line, "## ")
			if strings.EqualFold(trimmed, name) || strings.EqualFold(line, name) {
				// Save previous section.
				if currentSection != "" {
					sections[currentSection] = strings.TrimSpace(strings.Join(currentContent, "\n"))
				}
				currentSection = name
				if name == "Examples" {
					currentSection = "Example"
				}
				currentContent = nil
				found = true
				break
			}
		}
		if !found && currentSection != "" {
			currentContent = append(currentContent, line)
		}
	}

	// Save last section.
	if currentSection != "" {
		sections[currentSection] = strings.TrimSpace(strings.Join(currentContent, "\n"))
	}

	return sections
}
