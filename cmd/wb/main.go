package main

import (
	"fmt"
	"os"
	"time"
)

type globalFlags struct {
	format   string
	mode     string
	compact  bool
	start    time.Time
	warnings []string
}

const (
	modeDefault     = "default"
	modeAgent       = "agent"
	schemaVersionV1 = "wb.v1"
)

func main() {
	os.Exit(run(os.Args[1:]))
}

func run(args []string) int {
	globals := globalFlags{
		format: FormatJSON,
		mode:   modeDefault,
		start:  time.Now(),
	}

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
		case "--mode":
			if i+1 >= len(args) {
				writeError("", errUsage("--mode requires a value"), globals)
				return ExitUsage
			}
			globals.mode = args[i+1]
			i++
		case "--compact":
			globals.compact = true
		default:
			remaining = append(remaining, args[i])
		}
	}

	// Validate mode.
	switch globals.mode {
	case modeDefault, modeAgent:
		// ok
	default:
		writeError("", &ErrorInfo{
			Code:    ErrCodeUsage,
			Message: fmt.Sprintf("unknown mode %q", globals.mode),
			Hint:    "Supported modes: default, agent.",
		}, globals)
		return ExitUsage
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

	// Agent mode always emits JSON envelopes.
	if globals.mode == modeAgent && globals.format != FormatJSON {
		globals.warnings = append(globals.warnings, "agent mode forces --format json")
		globals.format = FormatJSON
	}

	if len(remaining) == 0 {
		if globals.mode == modeAgent {
			writeError("", errUsage("command required"), globals)
			return ExitUsage
		}
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
	resp := successResponse(command, data, buildMeta(command, globals))
	writeResponse(resp, globals, false)
}

func writeError(command string, ei *ErrorInfo, globals globalFlags) {
	resp := errorResponse(command, ei, buildMeta(command, globals))
	writeResponse(resp, globals, true)
}

func writeResponse(resp *Response, globals globalFlags, toStderr bool) {
	out, err := marshalJSON(resp, globals.compact)
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		return
	}
	if toStderr && globals.mode != modeAgent {
		fmt.Fprintln(os.Stderr, string(out))
		return
	}
	fmt.Println(string(out))
}

func buildMeta(command string, globals globalFlags) *responseMeta {
	meta := &responseMeta{
		SchemaVersion:         schemaVersionV1,
		ToolVersion:           version,
		ElapsedMS:             time.Since(globals.start).Milliseconds(),
		Mode:                  globals.mode,
		Warnings:              globals.warnings,
		NextSuggestedCommands: nextSuggestedCommands(command),
	}
	return meta
}

func nextSuggestedCommands(command string) []string {
	switch command {
	case "info":
		return []string{"wb read <file>", "wb dep <file>"}
	case "read":
		return []string{"wb calc <file>", "wb edit --patch '[...]' <file>"}
	case "edit":
		return []string{"wb read <file>", "wb calc <file>"}
	case "create":
		return []string{"wb info <file>", "wb read <file>"}
	case "calc":
		return []string{"wb read <file>", "wb dep <file>"}
	case "dep":
		return []string{"wb read <file>", "wb formula list"}
	case "formula":
		return []string{"wb formula list"}
	case "version":
		return []string{"wb help"}
	default:
		return []string{"wb help", "wb version"}
	}
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
  --mode <default|agent>         Output contract mode (default: default)
  --compact                      Emit compact JSON (no indentation)

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
