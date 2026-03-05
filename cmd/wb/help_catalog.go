package main

import (
	"fmt"
	"io"
	"os"
	"strings"
)

type toolSpec struct {
	Name        string        `json:"name"`
	Summary     string        `json:"summary"`
	Usage       string        `json:"usage"`
	Modes       []modeSpec    `json:"modes,omitempty"`
	GlobalFlags []flagSpec    `json:"global_flags,omitempty"`
	Commands    []commandSpec `json:"commands,omitempty"`
}

type modeSpec struct {
	Name        string   `json:"name"`
	Description string   `json:"description"`
	Effects     []string `json:"effects,omitempty"`
}

type flagSpec struct {
	Name          string   `json:"name"`
	ValueName     string   `json:"value_name,omitempty"`
	Description   string   `json:"description"`
	Default       string   `json:"default,omitempty"`
	AllowedValues []string `json:"allowed_values,omitempty"`
	Repeatable    bool     `json:"repeatable,omitempty"`
}

type commandSpec struct {
	Name             string        `json:"name"`
	Path             []string      `json:"path"`
	Summary          string        `json:"summary"`
	Description      string        `json:"description,omitempty"`
	Usage            string        `json:"usage"`
	Aliases          []string      `json:"aliases,omitempty"`
	RequiresFile     bool          `json:"requires_file,omitempty"`
	ReadsStdin       bool          `json:"reads_stdin,omitempty"`
	StdinJSONKind    string        `json:"stdin_json_kind,omitempty"`
	SupportedFormats []string      `json:"supported_formats,omitempty"`
	Flags            []flagSpec    `json:"flags,omitempty"`
	Notes            []string      `json:"notes,omitempty"`
	Examples         []string      `json:"examples,omitempty"`
	Subcommands      []commandSpec `json:"subcommands,omitempty"`
}

type helpData struct {
	Topic       string       `json:"topic"`
	Tool        *toolSpec    `json:"tool,omitempty"`
	Command     *commandSpec `json:"command,omitempty"`
	GlobalFlags []flagSpec   `json:"global_flags,omitempty"`
}

func wbToolSpec() toolSpec {
	return toolSpec{
		Name:    "wb",
		Summary: "Read, write, and inspect Excel .xlsx workbooks.",
		Usage:   "wb <command> [flags] <file>",
		Modes: []modeSpec{
			{
				Name:        modeDefault,
				Description: "Human-oriented mode. Successful JSON responses go to stdout; errors and help text go to stderr.",
			},
			{
				Name:        modeAgent,
				Description: "Agent-oriented mode. Success, errors, and help all return JSON envelopes on stdout.",
				Effects: []string{
					"Forces --format json.",
					"Returns structured help instead of plain text.",
				},
			},
		},
		GlobalFlags: []flagSpec{
			{
				Name:          "--format",
				ValueName:     "json|markdown|csv",
				Description:   "Output format. Only read supports markdown and csv output; agent mode always forces json.",
				Default:       FormatJSON,
				AllowedValues: []string{FormatJSON, FormatMarkdown, FormatCSV},
			},
			{
				Name:          "--mode",
				ValueName:     "default|agent",
				Description:   "Output contract mode.",
				Default:       modeDefault,
				AllowedValues: []string{modeDefault, modeAgent},
			},
			{
				Name:        "--compact",
				Description: "Emit compact JSON with no indentation.",
			},
		},
		Commands: []commandSpec{
			{
				Name:             "info",
				Path:             []string{"info"},
				Summary:          "Show sheet metadata (dimensions, cell counts)",
				Description:      "Show sheet metadata including dimensions, cell counts, and formula presence.",
				Usage:            "wb info [flags] <file>",
				RequiresFile:     true,
				SupportedFormats: []string{FormatJSON},
				Flags: []flagSpec{
					{
						Name:        "--sheet",
						ValueName:   "name",
						Description: "Show only the named sheet.",
						Default:     "all sheets",
					},
				},
				Examples: []string{
					"wb info data.xlsx",
					"wb info --sheet Sheet1 data.xlsx",
				},
			},
			{
				Name:             "read",
				Path:             []string{"read"},
				Summary:          "Read cell data for a range or full sheet",
				Description:      "Read cell data from a workbook. Returns stored or cached values.",
				Usage:            "wb read [flags] <file>",
				RequiresFile:     true,
				SupportedFormats: []string{FormatJSON, FormatMarkdown, FormatCSV},
				Flags: []flagSpec{
					{Name: "--sheet", ValueName: "name", Description: "Read from the named sheet.", Default: "first sheet"},
					{Name: "--all-sheets", Description: "Read all sheets. Mutually exclusive with --sheet."},
					{Name: "--range", ValueName: "A1:D10", Description: "Read a specific range.", Default: "full used range"},
					{Name: "--limit", ValueName: "N", Description: "Limit output to the first N data rows."},
					{Name: "--head", ValueName: "N", Description: "Alias for --limit."},
					{Name: "--where", ValueName: "expr", Description: "Filter rows with column=value style expressions.", Repeatable: true},
					{Name: "--include-formulas", Description: "Include formula strings in output."},
					{Name: "--include-styles", Description: "Include style objects in output."},
					{Name: "--style-summary", Description: "Include a human-readable style summary per cell."},
					{Name: "--headers", Description: "Treat the first row as headers."},
					{Name: "--no-dates", Description: "Disable date detection; show raw numbers."},
				},
				Notes: []string{
					"Use 'calc' to force formula recalculation instead of reading cached values.",
					"Date cells are automatically detected and returned with type \"date\" plus a \"formatted\" ISO 8601 string.",
				},
				Examples: []string{
					"wb read data.xlsx",
					"wb read --range A1:C10 data.xlsx",
					"wb read --headers --include-formulas data.xlsx",
					"wb read --format csv --headers data.xlsx",
					"wb read --limit 5 --headers data.xlsx",
					"wb read --all-sheets --format markdown data.xlsx",
					"wb read --headers --where \"Status=Failed\" data.xlsx",
					"wb read --style-summary data.xlsx",
				},
			},
			{
				Name:             "edit",
				Path:             []string{"edit"},
				Summary:          "Apply JSON patch array of cell changes",
				Description:      "Apply JSON patch operations to an existing workbook.",
				Usage:            "wb edit [flags] <file>",
				RequiresFile:     true,
				ReadsStdin:       true,
				StdinJSONKind:    "patch_op[]",
				SupportedFormats: []string{FormatJSON},
				Flags: []flagSpec{
					{Name: "--patch", ValueName: "json", Description: "Patch JSON array. If omitted, patch JSON is read from stdin."},
					{Name: "--sheet", ValueName: "name", Description: "Default sheet for operations.", Default: "first sheet"},
					{Name: "--output", ValueName: "path", Description: "Write to a different file.", Default: "overwrite input"},
					{Name: "--dry-run", Description: "Report changes without saving."},
					{Name: "--validate-only", Description: "Validate and apply in-memory only; never saves."},
					{Name: "--atomic", Description: "Save only if all operations succeed.", Default: "true"},
					{Name: "--no-atomic", Description: "Allow partial saves when operations fail."},
					{Name: "--plan", Description: "Include a normalized operation plan in output."},
				},
				Notes: []string{
					"Patch JSON can be passed with --patch or via stdin.",
					"Setting cell values does not auto-expand formula ranges. If you add data beyond a range like SUM(B2:B3), update the formula separately.",
				},
				Examples: []string{
					"wb edit --patch '[{\"cell\":\"A1\",\"value\":\"updated\"}]' data.xlsx",
					"echo '[{\"cell\":\"B1\",\"formula\":\"SUM(A1:A10)\"}]' | wb edit data.xlsx",
					"wb edit --dry-run --patch '[{\"cell\":\"A1\",\"clear\":true}]' data.xlsx",
				},
			},
			{
				Name:             "create",
				Path:             []string{"create"},
				Summary:          "Create new workbook from JSON spec",
				Description:      "Create a new workbook from a JSON spec.",
				Usage:            "wb create [flags] <file>",
				RequiresFile:     true,
				ReadsStdin:       true,
				StdinJSONKind:    "create_spec",
				SupportedFormats: []string{FormatJSON},
				Flags: []flagSpec{
					{Name: "--spec", ValueName: "json", Description: "Spec JSON. If omitted, spec JSON is read from stdin."},
				},
				Notes: []string{
					"Unknown JSON fields are rejected.",
					"The spec supports sheets, cells, and row-oriented data blocks.",
				},
				Examples: []string{
					"wb create --spec '{\"sheets\":[\"S1\"],\"cells\":[{\"cell\":\"A1\",\"value\":\"hello\"}]}' out.xlsx",
					"echo '{\"rows\":[{\"start\":\"A1\",\"data\":[[\"a\",\"b\"],[1,2]]}]}' | wb create out.xlsx",
				},
			},
			{
				Name:             "calc",
				Path:             []string{"calc"},
				Summary:          "Force recalculation and return results",
				Description:      "Force recalculation of all formulas and return evaluated results.",
				Usage:            "wb calc [flags] <file>",
				RequiresFile:     true,
				SupportedFormats: []string{FormatJSON},
				Flags: []flagSpec{
					{Name: "--sheet", ValueName: "name", Description: "Recalculate and show the named sheet.", Default: "first sheet"},
					{Name: "--range", ValueName: "A1:D10", Description: "Return results for a specific range.", Default: "full used range"},
					{Name: "--output", ValueName: "path", Description: "Save the recalculated workbook to a file."},
					{Name: "--no-dates", Description: "Disable date detection; show raw numbers."},
				},
				Notes: []string{
					"Unlike 'read', calc evaluates formulas instead of returning cached values.",
				},
				Examples: []string{
					"wb calc data.xlsx",
					"wb calc --range A1:C10 data.xlsx",
					"wb calc --output recalculated.xlsx data.xlsx",
				},
			},
			{
				Name:             "dep",
				Path:             []string{"dep"},
				Summary:          "Show cell dependency graph",
				Description:      "Show precedents and dependents for a cell, range, or all formulas on a sheet.",
				Usage:            "wb dep [flags] <file>",
				RequiresFile:     true,
				SupportedFormats: []string{FormatJSON},
				Flags: []flagSpec{
					{Name: "--cell", ValueName: "A1", Description: "Show dependencies for a single cell."},
					{Name: "--range", ValueName: "A1:D10", Description: "Show dependencies for all formula cells in a range."},
					{Name: "--sheet", ValueName: "name", Description: "Target sheet.", Default: "first sheet"},
					{Name: "--direction", ValueName: "precedents|dependents|both", Description: "Which edges to include.", Default: "both", AllowedValues: []string{"precedents", "dependents", "both"}},
					{Name: "--depth", ValueName: "N", Description: "Transitive depth. Use -1 for the full reachable graph.", Default: "1"},
				},
				Notes: []string{
					"When neither --cell nor --range is provided, wb discovers all formula cells on the target sheet.",
					"Dependents are searched across all sheets.",
				},
				Examples: []string{
					"wb dep data.xlsx",
					"wb dep --sheet Sheet2 data.xlsx",
					"wb dep --cell A1 data.xlsx",
					"wb dep --cell A1 --depth -1 data.xlsx",
					"wb dep --direction dependents --cell A1 data.xlsx",
				},
			},
			{
				Name:        "formula",
				Path:        []string{"formula"},
				Summary:     "Formula-related subcommands",
				Description: "Formula-related subcommands.",
				Usage:       "wb formula <subcommand>",
				Subcommands: []commandSpec{
					{
						Name:             "list",
						Path:             []string{"formula", "list"},
						Summary:          "List all registered formula functions",
						Description:      "List all registered formula functions.",
						Usage:            "wb formula list",
						SupportedFormats: []string{FormatJSON},
						Examples: []string{
							"wb formula list",
						},
					},
				},
				Examples: []string{
					"wb formula list",
				},
			},
			{
				Name:             "capabilities",
				Path:             []string{"capabilities"},
				Summary:          "Show machine-readable CLI metadata",
				Description:      "Return structured metadata for commands, flags, modes, and agent-mode behavior.",
				Usage:            "wb capabilities",
				SupportedFormats: []string{FormatJSON},
				Examples: []string{
					"wb capabilities",
					"wb --mode agent help read",
				},
				Notes: []string{
					"Use this command to discover commands and flags without scraping text help.",
				},
			},
			{
				Name:        "help",
				Path:        []string{"help"},
				Summary:     "Show help for a command",
				Description: "Show human-readable help in default mode or structured JSON help in agent mode.",
				Usage:       "wb help [command]",
				Aliases:     []string{"--help", "-h"},
				SupportedFormats: []string{
					FormatJSON,
				},
				Examples: []string{
					"wb help",
					"wb help read",
					"wb --mode agent help read",
				},
				Notes: []string{
					"In agent mode, help responses are returned as JSON envelopes on stdout.",
				},
			},
			{
				Name:             "version",
				Path:             []string{"version"},
				Summary:          "Print version info",
				Description:      "Print version info.",
				Usage:            "wb version",
				SupportedFormats: []string{FormatJSON},
				Examples: []string{
					"wb version",
				},
			},
		},
	}
}

func writeHelpTopic(path []string, globals globalFlags) int {
	spec := wbToolSpec()
	if globals.mode == modeAgent {
		if len(path) == 0 {
			writeSuccess("help", helpData{
				Topic: "overview",
				Tool:  &spec,
			}, globals)
			return ExitSuccess
		}

		cmd, ok := lookupCommandSpec(spec.Commands, path)
		if !ok {
			writeError("help", errUsage(fmt.Sprintf("unknown command %q", strings.Join(path, " "))), globals)
			return ExitUsage
		}

		writeSuccess("help", helpData{
			Topic:       "command",
			Command:     &cmd,
			GlobalFlags: spec.GlobalFlags,
		}, globals)
		return ExitSuccess
	}

	if len(path) == 0 {
		printUsage()
		return ExitSuccess
	}

	cmd, ok := lookupCommandSpec(spec.Commands, path)
	if !ok {
		writeError("help", errUsage(fmt.Sprintf("unknown command %q", strings.Join(path, " "))), globals)
		return ExitUsage
	}

	renderCommandHelp(os.Stderr, cmd, spec.GlobalFlags)
	return ExitSuccess
}

func lookupCommandSpec(commands []commandSpec, path []string) (commandSpec, bool) {
	if len(path) == 0 {
		return commandSpec{}, false
	}

	for _, cmd := range commands {
		if cmd.Name != path[0] {
			continue
		}
		if len(path) == 1 {
			return cmd, true
		}
		return lookupCommandSpec(cmd.Subcommands, path[1:])
	}

	return commandSpec{}, false
}

func renderToolHelp(w io.Writer, spec toolSpec) {
	fmt.Fprintf(w, "Usage: %s\n\n", spec.Usage)
	if spec.Summary != "" {
		fmt.Fprintf(w, "%s\n", spec.Summary)
	}

	if len(spec.Commands) > 0 {
		fmt.Fprintln(w, "\nCommands:")
		for _, cmd := range spec.Commands {
			fmt.Fprintf(w, "  %-12s %s\n", cmd.Name, cmd.Summary)
		}
	}

	if len(spec.GlobalFlags) > 0 {
		fmt.Fprintln(w, "\nGlobal flags:")
		renderFlags(w, spec.GlobalFlags)
	}

	fmt.Fprintln(w, "\nRun 'wb <command> --help' for detailed command usage.")
}

func renderCommandHelp(w io.Writer, spec commandSpec, globalFlags []flagSpec) {
	fmt.Fprintf(w, "Usage: %s\n\n", spec.Usage)

	desc := spec.Description
	if desc == "" {
		desc = spec.Summary
	}
	if desc != "" {
		fmt.Fprintf(w, "%s\n", desc)
	}

	if len(spec.Subcommands) > 0 {
		fmt.Fprintln(w, "\nSubcommands:")
		for _, sub := range spec.Subcommands {
			fmt.Fprintf(w, "  %-12s %s\n", sub.Name, sub.Summary)
		}
	}

	if len(spec.Flags) > 0 {
		fmt.Fprintln(w, "\nFlags:")
		renderFlags(w, spec.Flags)
	}

	if len(globalFlags) > 0 {
		fmt.Fprintln(w, "\nGlobal flags:")
		renderFlags(w, globalFlags)
	}

	if len(spec.Notes) > 0 {
		fmt.Fprintln(w, "\nNotes:")
		for _, note := range spec.Notes {
			fmt.Fprintf(w, "  %s\n", note)
		}
	}

	if len(spec.Examples) > 0 {
		fmt.Fprintln(w, "\nExamples:")
		for _, example := range spec.Examples {
			fmt.Fprintf(w, "  %s\n", example)
		}
	}
}

func renderFlags(w io.Writer, flags []flagSpec) {
	width := 0
	labels := make([]string, len(flags))
	for i, flag := range flags {
		label := flag.Name
		if flag.ValueName != "" {
			label += " <" + flag.ValueName + ">"
		}
		labels[i] = label
		if len(label) > width {
			width = len(label)
		}
	}

	for i, flag := range flags {
		desc := flag.Description
		if flag.Default != "" {
			desc += " (default: " + flag.Default + ")"
		}
		fmt.Fprintf(w, "  %-*s %s\n", width, labels[i], desc)
	}
}
