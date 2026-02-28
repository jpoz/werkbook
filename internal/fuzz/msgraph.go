package fuzz

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"math/rand"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// MSGraphConfig holds configuration for the MS Graph evaluator.
type MSGraphConfig struct {
	TenantID     string
	ClientID     string
	ClientSecret string
	TokenCache   string // path to token cache file
}

// tokenCache stores the OAuth tokens on disk.
type tokenCache struct {
	AccessToken  string    `json:"access_token"`
	RefreshToken string    `json:"refresh_token"`
	ExpiresAt    time.Time `json:"expires_at"`
}

// MSGraphEvaluator implements Evaluator using Excel Online via the MS Graph API.
type MSGraphEvaluator struct {
	config MSGraphConfig
	token  *tokenCache
	client *http.Client
}

// NewMSGraphEvaluator creates a new MSGraphEvaluator from environment variables.
func NewMSGraphEvaluator() (*MSGraphEvaluator, error) {
	tenantID := os.Getenv("MSGRAPH_TENANT_ID")
	if tenantID == "" {
		return nil, fmt.Errorf("MSGRAPH_TENANT_ID environment variable is required for excel oracle")
	}
	clientID := os.Getenv("MSGRAPH_CLIENT_ID")
	if clientID == "" {
		return nil, fmt.Errorf("MSGRAPH_CLIENT_ID environment variable is required for excel oracle")
	}
	clientSecret := os.Getenv("MSGRAPH_CLIENT_SECRET")
	tokenCachePath := os.Getenv("MSGRAPH_TOKEN_CACHE")
	if tokenCachePath == "" {
		home, err := os.UserHomeDir()
		if err != nil {
			return nil, fmt.Errorf("get home dir: %w", err)
		}
		tokenCachePath = filepath.Join(home, ".werkbook", "msgraph_token.json")
	}

	e := &MSGraphEvaluator{
		config: MSGraphConfig{
			TenantID:     tenantID,
			ClientID:     clientID,
			ClientSecret: clientSecret,
			TokenCache:   tokenCachePath,
		},
		client: &http.Client{Timeout: 120 * time.Second},
	}

	if err := e.loadToken(); err != nil {
		return nil, fmt.Errorf("load token: %w (run 'go run ./cmd/msgraph-setup' first)", err)
	}

	return e, nil
}

// Name returns "excel".
func (e *MSGraphEvaluator) Name() string { return "excel" }

// ExcludedFunctions returns the Excel Online exclusion list.
func (e *MSGraphEvaluator) ExcludedFunctions() map[string]bool {
	return ExcelOnlineExcludedFunctions
}

// Eval uploads the XLSX to OneDrive, creates a calculation session,
// reads the evaluated values, and returns results for the given checks.
func (e *MSGraphEvaluator) Eval(xlsxPath string, checks []CheckSpec) ([]CellResult, error) {
	if err := e.ensureValidToken(); err != nil {
		return nil, fmt.Errorf("refresh token: %w", err)
	}

	// Upload file.
	itemID, err := e.uploadFile(xlsxPath)
	if err != nil {
		return nil, fmt.Errorf("upload file: %w", err)
	}
	defer e.deleteFile(itemID) //nolint:errcheck

	// Create a non-persistent session to trigger calculation.
	sessionID, err := e.createSession(itemID)
	if err != nil {
		return nil, fmt.Errorf("create session: %w", err)
	}
	defer e.closeSession(itemID, sessionID) //nolint:errcheck

	// Collect all unique sheet names needed.
	sheetValues := make(map[string]map[string]string)

	var results []CellResult
	for _, check := range checks {
		sheet, cellRef := ParseCheckRef(check.Ref)
		if sheet == "" {
			sheet = "Sheet1"
		}

		// Parse sheet values on demand.
		if _, ok := sheetValues[sheet]; !ok {
			vals, err := e.readUsedRange(itemID, sessionID, sheet)
			if err != nil {
				return nil, fmt.Errorf("read sheet %q: %w", sheet, err)
			}
			sheetValues[sheet] = vals
		}

		val, ok := sheetValues[sheet][cellRef]
		if !ok {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: "",
				Type:  "empty",
			})
		} else {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: val,
				Type:  guessExcelType(val),
			})
		}
	}

	return results, nil
}

// --- Token management ---

func (e *MSGraphEvaluator) loadToken() error {
	data, err := os.ReadFile(e.config.TokenCache)
	if err != nil {
		return fmt.Errorf("read token cache %s: %w", e.config.TokenCache, err)
	}
	var tc tokenCache
	if err := json.Unmarshal(data, &tc); err != nil {
		return fmt.Errorf("parse token cache: %w", err)
	}
	if tc.RefreshToken == "" {
		return fmt.Errorf("token cache has no refresh token")
	}
	e.token = &tc
	return nil
}

func (e *MSGraphEvaluator) saveToken() error {
	dir := filepath.Dir(e.config.TokenCache)
	if err := os.MkdirAll(dir, 0700); err != nil {
		return fmt.Errorf("create token cache dir: %w", err)
	}
	data, err := json.MarshalIndent(e.token, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(e.config.TokenCache, data, 0600)
}

func (e *MSGraphEvaluator) ensureValidToken() error {
	if e.token == nil {
		return fmt.Errorf("no token loaded")
	}
	// Refresh if expired or within 5 minutes of expiry.
	if time.Now().After(e.token.ExpiresAt.Add(-5 * time.Minute)) {
		return e.refreshToken()
	}
	return nil
}

func (e *MSGraphEvaluator) refreshToken() error {
	endpoint := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", e.config.TenantID)

	form := url.Values{
		"client_id":     {e.config.ClientID},
		"grant_type":    {"refresh_token"},
		"refresh_token": {e.token.RefreshToken},
		"scope":         {"Files.ReadWrite User.Read offline_access"},
	}
	if e.config.ClientSecret != "" {
		form.Set("client_secret", e.config.ClientSecret)
	}

	resp, err := e.client.PostForm(endpoint, form)
	if err != nil {
		return fmt.Errorf("token refresh request: %w", err)
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("token refresh failed (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 500))
	}

	var tokenResp struct {
		AccessToken  string `json:"access_token"`
		RefreshToken string `json:"refresh_token"`
		ExpiresIn    int    `json:"expires_in"`
	}
	if err := json.Unmarshal(body, &tokenResp); err != nil {
		return fmt.Errorf("parse token response: %w", err)
	}

	e.token.AccessToken = tokenResp.AccessToken
	if tokenResp.RefreshToken != "" {
		e.token.RefreshToken = tokenResp.RefreshToken
	}
	e.token.ExpiresAt = time.Now().Add(time.Duration(tokenResp.ExpiresIn) * time.Second)

	return e.saveToken()
}

// --- Device code flow (used by setup command) ---

// DeviceCodeResponse holds the response from the device code endpoint.
type DeviceCodeResponse struct {
	DeviceCode      string `json:"device_code"`
	UserCode        string `json:"user_code"`
	VerificationURI string `json:"verification_uri"`
	ExpiresIn       int    `json:"expires_in"`
	Interval        int    `json:"interval"`
	Message         string `json:"message"`
}

// StartDeviceCodeFlow initiates the device code flow and returns the response.
func StartDeviceCodeFlow(tenantID, clientID string) (*DeviceCodeResponse, error) {
	endpoint := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/devicecode", tenantID)

	form := url.Values{
		"client_id": {clientID},
		"scope":     {"Files.ReadWrite User.Read offline_access"},
	}

	client := &http.Client{Timeout: 30 * time.Second}
	resp, err := client.PostForm(endpoint, form)
	if err != nil {
		return nil, fmt.Errorf("device code request: %w", err)
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("device code request failed (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 500))
	}

	var dcResp DeviceCodeResponse
	if err := json.Unmarshal(body, &dcResp); err != nil {
		return nil, fmt.Errorf("parse device code response: %w", err)
	}

	return &dcResp, nil
}

// PollForToken polls the token endpoint until the user authorizes or the device code expires.
func PollForToken(tenantID, clientID, clientSecret, deviceCode string, interval int) (*tokenCache, error) {
	endpoint := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", tenantID)

	if interval < 1 {
		interval = 5
	}

	client := &http.Client{Timeout: 30 * time.Second}

	for {
		time.Sleep(time.Duration(interval) * time.Second)

		form := url.Values{
			"client_id":   {clientID},
			"grant_type":  {"urn:ietf:params:oauth:grant-type:device_code"},
			"device_code": {deviceCode},
		}
		if clientSecret != "" {
			form.Set("client_secret", clientSecret)
		}

		resp, err := client.PostForm(endpoint, form)
		if err != nil {
			return nil, fmt.Errorf("token poll request: %w", err)
		}

		body, _ := io.ReadAll(resp.Body)
		resp.Body.Close()

		if resp.StatusCode == http.StatusOK {
			var tokenResp struct {
				AccessToken  string `json:"access_token"`
				RefreshToken string `json:"refresh_token"`
				ExpiresIn    int    `json:"expires_in"`
			}
			if err := json.Unmarshal(body, &tokenResp); err != nil {
				return nil, fmt.Errorf("parse token response: %w", err)
			}
			return &tokenCache{
				AccessToken:  tokenResp.AccessToken,
				RefreshToken: tokenResp.RefreshToken,
				ExpiresAt:    time.Now().Add(time.Duration(tokenResp.ExpiresIn) * time.Second),
			}, nil
		}

		// Check error type.
		var errResp struct {
			Error string `json:"error"`
		}
		if err := json.Unmarshal(body, &errResp); err == nil {
			switch errResp.Error {
			case "authorization_pending":
				continue // Keep polling.
			case "slow_down":
				interval += 5
				continue
			case "expired_token":
				return nil, fmt.Errorf("device code expired — please retry")
			case "authorization_declined":
				return nil, fmt.Errorf("authorization was declined")
			default:
				return nil, fmt.Errorf("token poll error: %s (body: %s)", errResp.Error, truncateBytes(body, 300))
			}
		}

		return nil, fmt.Errorf("unexpected token poll response (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 300))
	}
}

// SaveTokenCache writes a token cache to the given path.
func SaveTokenCache(path string, tc *tokenCache) error {
	dir := filepath.Dir(path)
	if err := os.MkdirAll(dir, 0700); err != nil {
		return fmt.Errorf("create token cache dir: %w", err)
	}
	data, err := json.MarshalIndent(tc, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0600)
}

// --- MS Graph API operations ---

func (e *MSGraphEvaluator) uploadFile(xlsxPath string) (string, error) {
	data, err := os.ReadFile(xlsxPath)
	if err != nil {
		return "", fmt.Errorf("read xlsx: %w", err)
	}

	// Use a unique filename to avoid conflicts.
	uniqueName := fmt.Sprintf("werkbook-fuzz-%d-%d.xlsx", time.Now().UnixNano(), rand.Intn(10000))
	uploadURL := fmt.Sprintf("https://graph.microsoft.com/v1.0/me/drive/root:/werkbook-fuzz/%s:/content", uniqueName)

	resp, err := e.doRequest("PUT", uploadURL, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", bytes.NewReader(data))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK && resp.StatusCode != http.StatusCreated {
		return "", fmt.Errorf("upload failed (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 500))
	}

	var item struct {
		ID string `json:"id"`
	}
	if err := json.Unmarshal(body, &item); err != nil {
		return "", fmt.Errorf("parse upload response: %w", err)
	}

	return item.ID, nil
}

func (e *MSGraphEvaluator) createSession(itemID string) (string, error) {
	sessionURL := fmt.Sprintf("https://graph.microsoft.com/v1.0/me/drive/items/%s/workbook/createSession", itemID)

	payload := []byte(`{"persistChanges": false}`)
	resp, err := e.doRequest("POST", sessionURL, "application/json", bytes.NewReader(payload))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK && resp.StatusCode != http.StatusCreated {
		return "", fmt.Errorf("create session failed (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 500))
	}

	var session struct {
		ID string `json:"id"`
	}
	if err := json.Unmarshal(body, &session); err != nil {
		return "", fmt.Errorf("parse session response: %w", err)
	}

	return session.ID, nil
}

func (e *MSGraphEvaluator) readUsedRange(itemID, sessionID, sheetName string) (map[string]string, error) {
	encodedSheet := url.PathEscape(sheetName)
	rangeURL := fmt.Sprintf(
		"https://graph.microsoft.com/v1.0/me/drive/items/%s/workbook/worksheets/%s/usedRange",
		itemID, encodedSheet,
	)

	req, err := http.NewRequest("GET", rangeURL, nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+e.token.AccessToken)
	if sessionID != "" {
		req.Header.Set("workbook-session-id", sessionID)
	}

	resp, err := e.doRequestRaw(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("read used range failed (HTTP %d): %s", resp.StatusCode, truncateBytes(body, 500))
	}

	return parseUsedRangeResponse(body)
}

func (e *MSGraphEvaluator) closeSession(itemID, sessionID string) error {
	sessionURL := fmt.Sprintf("https://graph.microsoft.com/v1.0/me/drive/items/%s/workbook/closeSession", itemID)

	req, err := http.NewRequest("POST", sessionURL, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+e.token.AccessToken)
	req.Header.Set("workbook-session-id", sessionID)
	req.Header.Set("Content-Length", "0")

	resp, err := e.doRequestRaw(req)
	if err != nil {
		return err
	}
	resp.Body.Close()
	return nil
}

func (e *MSGraphEvaluator) deleteFile(itemID string) error {
	deleteURL := fmt.Sprintf("https://graph.microsoft.com/v1.0/me/drive/items/%s", itemID)

	req, err := http.NewRequest("DELETE", deleteURL, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+e.token.AccessToken)

	resp, err := e.doRequestRaw(req)
	if err != nil {
		return err
	}
	resp.Body.Close()
	return nil
}

// --- HTTP helpers with retry ---

func (e *MSGraphEvaluator) doRequest(method, urlStr, contentType string, body io.Reader) (*http.Response, error) {
	// Read body into buffer for retries.
	var bodyBytes []byte
	if body != nil {
		var err error
		bodyBytes, err = io.ReadAll(body)
		if err != nil {
			return nil, fmt.Errorf("read request body: %w", err)
		}
	}

	var lastErr error
	for attempt := 0; attempt < 5; attempt++ {
		var reqBody io.Reader
		if bodyBytes != nil {
			reqBody = bytes.NewReader(bodyBytes)
		}

		req, err := http.NewRequest(method, urlStr, reqBody)
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+e.token.AccessToken)
		if contentType != "" {
			req.Header.Set("Content-Type", contentType)
		}

		resp, err := e.client.Do(req)
		if err != nil {
			lastErr = err
			time.Sleep(backoff(attempt))
			continue
		}

		switch resp.StatusCode {
		case http.StatusTooManyRequests, http.StatusServiceUnavailable:
			resp.Body.Close()
			lastErr = fmt.Errorf("HTTP %d", resp.StatusCode)
			time.Sleep(backoff(attempt))
			continue
		case http.StatusUnauthorized:
			resp.Body.Close()
			if err := e.refreshToken(); err != nil {
				return nil, fmt.Errorf("auto-refresh token: %w", err)
			}
			lastErr = fmt.Errorf("HTTP 401 (retrying after token refresh)")
			continue
		}

		return resp, nil
	}

	return nil, fmt.Errorf("request failed after retries: %w", lastErr)
}

func (e *MSGraphEvaluator) doRequestRaw(req *http.Request) (*http.Response, error) {
	var lastErr error
	for attempt := 0; attempt < 5; attempt++ {
		resp, err := e.client.Do(req)
		if err != nil {
			lastErr = err
			time.Sleep(backoff(attempt))
			continue
		}

		switch resp.StatusCode {
		case http.StatusTooManyRequests, http.StatusServiceUnavailable:
			resp.Body.Close()
			lastErr = fmt.Errorf("HTTP %d", resp.StatusCode)
			time.Sleep(backoff(attempt))
			continue
		case http.StatusUnauthorized:
			resp.Body.Close()
			if err := e.refreshToken(); err != nil {
				return nil, fmt.Errorf("auto-refresh token: %w", err)
			}
			req.Header.Set("Authorization", "Bearer "+e.token.AccessToken)
			lastErr = fmt.Errorf("HTTP 401 (retrying after token refresh)")
			continue
		}

		return resp, nil
	}

	return nil, fmt.Errorf("request failed after retries: %w", lastErr)
}

func backoff(attempt int) time.Duration {
	base := time.Duration(1<<uint(attempt)) * time.Second
	if base > 30*time.Second {
		base = 30 * time.Second
	}
	// Add jitter.
	jitter := time.Duration(rand.Intn(1000)) * time.Millisecond
	return base + jitter
}

// --- Response parsing ---

// usedRangeResponse represents the MS Graph usedRange API response.
type usedRangeResponse struct {
	Address    string `json:"address"`
	ColumnIndex int   `json:"columnIndex"`
	RowIndex    int   `json:"rowIndex"`
	Values     [][]any `json:"values"`
}

// parseUsedRangeResponse extracts cell values from a usedRange API response.
// Returns a map of cell ref (e.g. "A1") to string value.
func parseUsedRangeResponse(body []byte) (map[string]string, error) {
	var rng usedRangeResponse
	if err := json.Unmarshal(body, &rng); err != nil {
		return nil, fmt.Errorf("parse used range: %w", err)
	}

	// Parse the address to find the starting cell.
	// Address format: "Sheet1!A1:D5" or "A1:D5"
	startCol, startRow := parseRangeStart(rng.Address)
	if startCol == 0 {
		startCol = rng.ColumnIndex + 1
	}
	if startRow == 0 {
		startRow = rng.RowIndex + 1
	}

	result := make(map[string]string)
	for rowIdx, row := range rng.Values {
		for colIdx, val := range row {
			strVal := formatGraphValue(val)
			if strVal == "" {
				continue
			}
			ref := fmt.Sprintf("%s%d", ColToLetter(startCol+colIdx), startRow+rowIdx)
			result[ref] = strVal
		}
	}

	return result, nil
}

// parseRangeStart extracts the starting column and row from a range address.
// e.g. "Sheet1!A1:D5" → (1, 1), "B2:C3" → (2, 2)
func parseRangeStart(address string) (col, row int) {
	// Strip sheet name if present.
	if idx := strings.Index(address, "!"); idx >= 0 {
		address = address[idx+1:]
	}
	// Take the start cell of the range.
	if idx := strings.Index(address, ":"); idx >= 0 {
		address = address[:idx]
	}
	// Parse the cell reference.
	colStr := ""
	rowStr := ""
	for _, ch := range address {
		if ch >= 'A' && ch <= 'Z' {
			colStr += string(ch)
		} else if ch >= '0' && ch <= '9' {
			rowStr += string(ch)
		}
	}
	if colStr == "" || rowStr == "" {
		return 0, 0
	}
	// Convert column letters to number.
	c := 0
	for _, ch := range colStr {
		c = c*26 + int(ch-'A') + 1
	}
	// Convert row string to number.
	r := 0
	for _, ch := range rowStr {
		r = r*10 + int(ch-'0')
	}
	return c, r
}

// formatGraphValue converts a value from the MS Graph API to a string.
// Graph returns: numbers as float64, booleans as bool, strings as string, errors as string.
func formatGraphValue(val any) string {
	if val == nil {
		return ""
	}
	switch v := val.(type) {
	case string:
		return v
	case float64:
		// Format without unnecessary trailing zeros.
		if v == float64(int64(v)) {
			return fmt.Sprintf("%d", int64(v))
		}
		return fmt.Sprintf("%g", v)
	case bool:
		if v {
			return "TRUE"
		}
		return "FALSE"
	default:
		return fmt.Sprintf("%v", v)
	}
}

// guessExcelType infers a type string from an Excel Online value.
func guessExcelType(val string) string {
	if val == "" {
		return "empty"
	}
	upper := strings.ToUpper(val)
	if upper == "TRUE" || upper == "FALSE" {
		return "bool"
	}
	if strings.HasPrefix(val, "#") {
		return "error"
	}
	return "number"
}

func truncateBytes(b []byte, maxLen int) string {
	s := string(b)
	if len(s) <= maxLen {
		return s
	}
	return s[:maxLen] + "..."
}

