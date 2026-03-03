package main

import (
	"fmt"
	"os"
)

type globalFlags struct {
	format string
}

func main() {
	os.Exit(run(os.Args[1:]))
}

func run(args []string) int {
	globals := globalFlags{format: FormatJSON}

	// Extract global flags from args.
	var remaining []string
	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--format":
			if i+1 >= len(args) {
				writeError("", errUsage("--format requires a value"), globals)
				return ExitUsage
			}
			globals.format = args[i+1]
			i++
		default:
			remaining = append(remaining, args[i])
		}
	}

	// Validate format.
	switch globals.format {
	case FormatJSON, FormatMarkdown, FormatCSV:
		// ok
	default:
		writeError("", &ErrorInfo{
			Code:    ErrCodeInvalidFormat,
			Message: fmt.Sprintf("unknown format %q", globals.format),
			Hint:    "Supported formats: json, markdown, csv.",
		}, globals)
		return ExitUsage
	}

	if len(remaining) == 0 {
		printUsage()
		return ExitUsage
	}

	command := remaining[0]
	cmdArgs := remaining[1:]

	switch command {
	case "info":
		return cmdInfo(cmdArgs, globals)
	case "read":
		return cmdRead(cmdArgs, globals)
	case "edit":
		return cmdEdit(cmdArgs, globals)
	case "create":
		return cmdCreate(cmdArgs, globals)
	case "calc":
		return cmdCalc(cmdArgs, globals)
	case "dep":
		return cmdDep(cmdArgs, globals)
	case "formula":
		return cmdFormula(cmdArgs, globals)
	case "version":
		return cmdVersion(globals)
	case "help", "--help", "-h":
		return cmdHelp(cmdArgs, globals)
	default:
		writeError("", errUsage(fmt.Sprintf("unknown command %q", command)), globals)
		return ExitUsage
	}
}

func writeSuccess(command string, data any, globals globalFlags) {
	resp := successResponse(command, data)
	out, err := marshalJSON(resp)
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		return
	}
	fmt.Println(string(out))
}

func writeError(command string, ei *ErrorInfo, globals globalFlags) {
	resp := errorResponse(command, ei)
	out, err := marshalJSON(resp)
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		return
	}
	fmt.Fprintln(os.Stderr, string(out))
}

func printUsage() {
	fmt.Fprintln(os.Stderr, `Usage: wb <command> [flags] <file>

Commands:
  info      Show sheet metadata (dimensions, cell counts)
  read      Read cell data for a range or full sheet
  edit      Apply JSON patch array of cell changes
  create    Create new workbook from JSON spec
  calc      Force recalculation and return results
  dep       Show cell dependency graph (precedents and dependents)
  formula   Formula-related subcommands (e.g. 'formula list')
  version   Print version info

Global flags:
  --format <json|markdown|csv>   Output format (default: json)

Run 'wb <command> --help' for detailed command usage.`)
}

// hasHelpFlag checks if --help is present in the args.
func hasHelpFlag(args []string) bool {
	for _, a := range args {
		if a == "--help" || a == "-h" {
			return true
		}
	}
	return false
}

// cmdHelp dispatches to the per-command help by injecting --help.
func cmdHelp(args []string, globals globalFlags) int {
	if len(args) == 0 {
		printUsage()
		return ExitSuccess
	}
	switch args[0] {
	case "info":
		return cmdInfo([]string{"--help"}, globals)
	case "read":
		return cmdRead([]string{"--help"}, globals)
	case "edit":
		return cmdEdit([]string{"--help"}, globals)
	case "create":
		return cmdCreate([]string{"--help"}, globals)
	case "calc":
		return cmdCalc([]string{"--help"}, globals)
	case "dep":
		return cmdDep([]string{"--help"}, globals)
	case "formula":
		return cmdFormula([]string{"--help"}, globals)
	case "version":
		return cmdVersion(globals)
	default:
		writeError("", errUsage(fmt.Sprintf("unknown command %q", args[0])), globals)
		return ExitUsage
	}
}
