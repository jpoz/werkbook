package main

import (
	"flag"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func main() {
	tenantID := flag.String("tenant-id", "", "Azure AD tenant ID (or set MSGRAPH_TENANT_ID)")
	clientID := flag.String("client-id", "", "Azure AD app registration client ID (or set MSGRAPH_CLIENT_ID)")
	clientSecret := flag.String("client-secret", "", "Azure AD client secret (or set MSGRAPH_CLIENT_SECRET)")
	tokenCache := flag.String("token-cache", "", "token cache path (default: ~/.werkbook/msgraph_token.json)")
	flag.Parse()

	// Resolve tenant ID.
	tid := *tenantID
	if tid == "" {
		tid = os.Getenv("MSGRAPH_TENANT_ID")
	}
	if tid == "" {
		fmt.Fprintf(os.Stderr, "Error: --tenant-id or MSGRAPH_TENANT_ID is required\n")
		flag.Usage()
		os.Exit(1)
	}

	// Resolve client ID.
	cid := *clientID
	if cid == "" {
		cid = os.Getenv("MSGRAPH_CLIENT_ID")
	}
	if cid == "" {
		fmt.Fprintf(os.Stderr, "Error: --client-id or MSGRAPH_CLIENT_ID is required\n")
		flag.Usage()
		os.Exit(1)
	}

	// Resolve client secret.
	csec := *clientSecret
	if csec == "" {
		csec = os.Getenv("MSGRAPH_CLIENT_SECRET")
	}

	// Resolve token cache path.
	cachePath := *tokenCache
	if cachePath == "" {
		cachePath = os.Getenv("MSGRAPH_TOKEN_CACHE")
	}
	if cachePath == "" {
		home, err := os.UserHomeDir()
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error getting home dir: %v\n", err)
			os.Exit(1)
		}
		cachePath = filepath.Join(home, ".werkbook", "msgraph_token.json")
	}

	fmt.Println("MS Graph Setup for Werkbook Fuzz Testing")
	fmt.Println("=========================================")
	fmt.Printf("Tenant ID:      %s\n", tid)
	fmt.Printf("Client ID:      %s\n", cid)
	if csec != "" {
		fmt.Printf("Client secret:  (set)\n")
	} else {
		fmt.Printf("Client secret:  (not set)\n")
	}
	fmt.Printf("Token cache:    %s\n", cachePath)
	fmt.Println()

	// Start device code flow.
	fmt.Println("Starting device code flow...")
	dcResp, err := fuzz.StartDeviceCodeFlow(tid, cid)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Println()
	fmt.Println(dcResp.Message)
	fmt.Println()

	// Open the verification URL in the default browser.
	if url := dcResp.VerificationURI; url != "" {
		fmt.Println("Opening browser...")
		var cmd *exec.Cmd
		switch runtime.GOOS {
		case "darwin":
			cmd = exec.Command("open", url)
		case "linux":
			cmd = exec.Command("xdg-open", url)
		case "windows":
			cmd = exec.Command("rundll32", "url.dll,FileProtocolHandler", url)
		}
		if cmd != nil {
			if err := cmd.Start(); err != nil {
				fmt.Fprintf(os.Stderr, "Could not open browser: %v\n", err)
			}
		}
	}

	// Poll for token.
	fmt.Println("Waiting for authorization...")
	tc, err := fuzz.PollForToken(tid, cid, csec, dcResp.DeviceCode, dcResp.Interval)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	// Save token.
	if err := fuzz.SaveTokenCache(cachePath, tc); err != nil {
		fmt.Fprintf(os.Stderr, "Error saving token: %v\n", err)
		os.Exit(1)
	}

	fmt.Println()
	fmt.Printf("Token saved to %s\n", cachePath)
	fmt.Println("Setup complete! You can now use --oracle excel with fuzz commands.")
}
