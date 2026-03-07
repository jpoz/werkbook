package main

import (
	"errors"
	"fmt"
	"regexp"
	"strings"

	"github.com/spf13/cobra"
	"github.com/spf13/pflag"
)

type cliState struct {
	exitCode int
}

type cliFlagError struct {
	command string
	message string
}

func (e *cliFlagError) Error() string {
	return e.message
}

var unknownCommandPattern = regexp.MustCompile(`^unknown command "([^"]+)" for "([^"]+)"$`)

func newRootCommand(globals globalFlags) (*cobra.Command, *cliState) {
	state := &cliState{exitCode: ExitSuccess}
	tool := wbToolSpec()

	root := &cobra.Command{
		Use:                tool.Name,
		Short:              tool.Summary,
		Long:               tool.Summary,
		SilenceErrors:      true,
		SilenceUsage:       true,
		TraverseChildren:   true,
		DisableSuggestions: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return nil
			}
			writeError("", errUsage(fmt.Sprintf("unknown command %q", args[0])), globals)
			state.exitCode = ExitUsage
			return nil
		},
		CompletionOptions: cobra.CompletionOptions{
			DisableDefaultCmd: true,
		},
	}

	addGlobalFlags(root.PersistentFlags())
	root.SetFlagErrorFunc(func(cmd *cobra.Command, err error) error {
		return &cliFlagError{
			command: commandName(commandPath(cmd)),
			message: normalizeFlagErrorMessage(err.Error()),
		}
	})
	root.SetHelpFunc(func(cmd *cobra.Command, args []string) {
		state.exitCode = writeHelpTopic(commandPath(cmd), globals)
	})
	root.SetHelpCommand(newHelpCommand(globals, state))

	root.AddCommand(
		newInfoCommand(globals, state),
		newReadCommand(globals, state),
		newEditCommand(globals, state),
		newCreateCommand(globals, state),
		newCalcCommand(globals, state),
		newDepCommand(globals, state),
		newFormulaCommand(globals, state),
		newCapabilitiesCommand(globals, state),
		newVersionCommand(globals, state),
	)

	return root, state
}

func addGlobalFlags(flags *pflag.FlagSet) {
	flags.String("format", FormatText, lookupGlobalFlagSpec("--format").Description)
	flags.String("mode", modeDefault, lookupGlobalFlagSpec("--mode").Description)
	flags.Bool("compact", false, lookupGlobalFlagSpec("--compact").Description)
}

func newInfoCommand(globals globalFlags, state *cliState) *cobra.Command {
	var sheet string

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--sheet", sheet)
			argv.addArgs(args)
			state.exitCode = cmdInfo(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("info"))
	cmd.Flags().StringVar(&sheet, "sheet", "", "Show only the named sheet.")
	return cmd
}

func newReadCommand(globals globalFlags, state *cliState) *cobra.Command {
	var sheet string
	var allSheets bool
	var rangeFlag string
	var limit string
	var where []string
	var includeFormulas bool
	var showFormulas bool
	var includeStyles bool
	var styleSummary bool
	var headers bool
	var noDates bool

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--sheet", sheet)
			argv.addBool("--all-sheets", allSheets)
			argv.addString("--range", rangeFlag)
			argv.addString("--limit", limit)
			argv.addStrings("--where", where)
			argv.addBool("--include-formulas", includeFormulas)
			argv.addBool("--show-formulas", showFormulas)
			argv.addBool("--include-styles", includeStyles)
			argv.addBool("--style-summary", styleSummary)
			argv.addBool("--headers", headers)
			argv.addBool("--no-dates", noDates)
			argv.addArgs(args)
			state.exitCode = cmdRead(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("read"))
	cmd.Flags().StringVar(&sheet, "sheet", "", "Read from the named sheet.")
	cmd.Flags().BoolVar(&allSheets, "all-sheets", false, "Read all sheets.")
	cmd.Flags().StringVar(&rangeFlag, "range", "", "Read a specific range.")
	cmd.Flags().StringVar(&limit, "limit", "", "Limit output to the first N data rows.")
	cmd.Flags().StringVar(&limit, "head", "", "Alias for --limit.")
	cmd.Flags().StringArrayVar(&where, "where", nil, "Filter rows with column=value style expressions.")
	cmd.Flags().BoolVar(&includeFormulas, "include-formulas", false, "Include formula strings in output.")
	cmd.Flags().BoolVar(&showFormulas, "show-formulas", false, "Display formula text instead of cached values.")
	cmd.Flags().BoolVar(&includeStyles, "include-styles", false, "Include style objects in output.")
	cmd.Flags().BoolVar(&styleSummary, "style-summary", false, "Include a human-readable style summary per cell.")
	cmd.Flags().BoolVar(&headers, "headers", false, "Treat the first row as headers.")
	cmd.Flags().BoolVar(&noDates, "no-dates", false, "Disable date detection; show raw numbers.")
	return cmd
}

func newEditCommand(globals globalFlags, state *cliState) *cobra.Command {
	var sheet string
	var patch string
	var output string
	var dryRun bool
	var validateOnly bool
	var atomic bool
	var noAtomic bool
	var plan bool

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--sheet", sheet)
			argv.addString("--patch", patch)
			argv.addString("--output", output)
			argv.addBool("--dry-run", dryRun)
			argv.addBool("--validate-only", validateOnly)
			if flagChanged(cmd, "atomic") {
				argv.addBool("--atomic", atomic)
			}
			if flagChanged(cmd, "no-atomic") {
				argv.addBool("--no-atomic", noAtomic)
			}
			argv.addBool("--plan", plan)
			argv.addArgs(args)
			state.exitCode = cmdEdit(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("edit"))
	cmd.Flags().StringVar(&sheet, "sheet", "", "Default sheet for operations.")
	cmd.Flags().StringVar(&patch, "patch", "", "Patch JSON array.")
	cmd.Flags().StringVar(&output, "output", "", "Write to a different file.")
	cmd.Flags().BoolVar(&dryRun, "dry-run", false, "Report changes without saving.")
	cmd.Flags().BoolVar(&validateOnly, "validate-only", false, "Validate and apply in-memory only.")
	cmd.Flags().BoolVar(&atomic, "atomic", false, "Save only if all operations succeed.")
	cmd.Flags().BoolVar(&noAtomic, "no-atomic", false, "Allow partial saves when operations fail.")
	cmd.Flags().BoolVar(&plan, "plan", false, "Include a normalized operation plan in output.")
	return cmd
}

func newCreateCommand(globals globalFlags, state *cliState) *cobra.Command {
	var specJSON string

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--spec", specJSON)
			argv.addArgs(args)
			state.exitCode = cmdCreate(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("create"))
	cmd.Flags().StringVar(&specJSON, "spec", "", "Spec JSON.")
	return cmd
}

func newCalcCommand(globals globalFlags, state *cliState) *cobra.Command {
	var sheet string
	var rangeFlag string
	var output string
	var noDates bool

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--sheet", sheet)
			argv.addString("--range", rangeFlag)
			argv.addString("--output", output)
			argv.addBool("--no-dates", noDates)
			argv.addArgs(args)
			state.exitCode = cmdCalc(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("calc"))
	cmd.Flags().StringVar(&sheet, "sheet", "", "Recalculate and show the named sheet.")
	cmd.Flags().StringVar(&rangeFlag, "range", "", "Return results for a specific range.")
	cmd.Flags().StringVar(&output, "output", "", "Save the recalculated workbook to a file.")
	cmd.Flags().BoolVar(&noDates, "no-dates", false, "Disable date detection; show raw numbers.")
	return cmd
}

func newDepCommand(globals globalFlags, state *cliState) *cobra.Command {
	var cell string
	var rangeFlag string
	var sheet string
	var direction string
	var depth string

	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			argv := new(argList)
			argv.addString("--cell", cell)
			argv.addString("--range", rangeFlag)
			argv.addString("--sheet", sheet)
			argv.addString("--direction", direction)
			argv.addString("--depth", depth)
			argv.addArgs(args)
			state.exitCode = cmdDep(argv.args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("dep"))
	cmd.Flags().StringVar(&cell, "cell", "", "Show dependencies for a single cell.")
	cmd.Flags().StringVar(&rangeFlag, "range", "", "Show dependencies for all formula cells in a range.")
	cmd.Flags().StringVar(&sheet, "sheet", "", "Target sheet.")
	cmd.Flags().StringVar(&direction, "direction", "", "Which edges to include.")
	cmd.Flags().StringVar(&depth, "depth", "", "Transitive depth.")
	return cmd
}

func newFormulaCommand(globals globalFlags, state *cliState) *cobra.Command {
	formulaCmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			state.exitCode = cmdFormula(args, globals)
			return nil
		},
	}
	applyCommandSpec(formulaCmd, mustLookupCommandSpec("formula"))

	listCmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			state.exitCode = cmdFormulaList(globals)
			return nil
		},
	}
	applyCommandSpec(listCmd, mustLookupCommandSpec("formula", "list"))
	formulaCmd.AddCommand(listCmd)

	return formulaCmd
}

func newCapabilitiesCommand(globals globalFlags, state *cliState) *cobra.Command {
	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			state.exitCode = cmdCapabilities(args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("capabilities"))
	return cmd
}

func newVersionCommand(globals globalFlags, state *cliState) *cobra.Command {
	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			state.exitCode = cmdVersion(args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("version"))
	return cmd
}

func newHelpCommand(globals globalFlags, state *cliState) *cobra.Command {
	cmd := &cobra.Command{
		RunE: func(cmd *cobra.Command, args []string) error {
			state.exitCode = cmdHelp(args, globals)
			return nil
		},
	}
	applyCommandSpec(cmd, mustLookupCommandSpec("help"))
	return cmd
}

func applyCommandSpec(cmd *cobra.Command, spec commandSpec) {
	use := strings.TrimPrefix(spec.Usage, "wb ")
	if use == "" {
		use = spec.Name
	}
	cmd.Use = use
	cmd.Short = spec.Summary
	if spec.Description != "" {
		cmd.Long = spec.Description
	} else {
		cmd.Long = spec.Summary
	}
	if len(spec.Examples) > 0 {
		cmd.Example = strings.Join(spec.Examples, "\n")
	}
}

func mustLookupCommandSpec(path ...string) commandSpec {
	spec, ok := lookupCommandSpec(wbToolSpec().Commands, path)
	if !ok {
		panic("missing command spec for " + strings.Join(path, " "))
	}
	return spec
}

func lookupGlobalFlagSpec(name string) flagSpec {
	for _, flag := range wbToolSpec().GlobalFlags {
		if flag.Name == name {
			return flag
		}
	}
	return flagSpec{Name: name}
}

func commandPath(cmd *cobra.Command) []string {
	var path []string
	for current := cmd; current != nil && current.Parent() != nil; current = current.Parent() {
		path = append(path, current.Name())
	}
	for i, j := 0, len(path)-1; i < j; i, j = i+1, j-1 {
		path[i], path[j] = path[j], path[i]
	}
	return path
}

func commandName(path []string) string {
	if len(path) == 0 {
		return ""
	}
	return path[len(path)-1]
}

func handleCLIError(err error, globals globalFlags) int {
	if err == nil || errors.Is(err, pflag.ErrHelp) {
		return ExitSuccess
	}

	var flagErr *cliFlagError
	if errors.As(err, &flagErr) {
		writeError(flagErr.command, errUsage(flagErr.message), globals)
		return ExitUsage
	}

	if command, message, ok := normalizeCommandErrorMessage(err.Error()); ok {
		writeError(command, errUsage(message), globals)
		return ExitUsage
	}

	writeError("", errInternal(err), globals)
	return ExitInternal
}

func normalizeFlagErrorMessage(message string) string {
	switch {
	case strings.HasPrefix(message, "flag needs an argument: "):
		name := strings.TrimPrefix(message, "flag needs an argument: ")
		return name + " requires a value"
	case strings.HasPrefix(message, "unknown shorthand flag: "):
		parts := strings.Split(message, " ")
		if len(parts) >= 4 {
			return "unknown flag: " + parts[3]
		}
	}
	return message
}

func normalizeCommandErrorMessage(message string) (string, string, bool) {
	match := unknownCommandPattern.FindStringSubmatch(strings.TrimSpace(message))
	if len(match) != 3 {
		return "", "", false
	}

	unknown := match[1]
	parentUse := match[2]
	parentPath := strings.Fields(parentUse)
	if len(parentPath) == 0 || parentPath[0] != "wb" {
		return "", fmt.Sprintf("unknown command %q", unknown), true
	}

	parentPath = parentPath[1:]
	if len(parentPath) == 0 {
		return "", fmt.Sprintf("unknown command %q", unknown), true
	}

	parentSpec, ok := lookupCommandSpec(wbToolSpec().Commands, parentPath)
	if !ok || len(parentSpec.Subcommands) == 0 {
		return commandName(parentPath), fmt.Sprintf("unknown command %q", unknown), true
	}

	var names []string
	for _, sub := range parentSpec.Subcommands {
		names = append(names, sub.Name)
	}
	return commandName(parentPath), fmt.Sprintf("unknown subcommand: %s. Available: %s", unknown, strings.Join(names, ", ")), true
}

func flagChanged(cmd *cobra.Command, name string) bool {
	flag := cmd.Flags().Lookup(name)
	return flag != nil && flag.Changed
}

type argList struct {
	args []string
}

func (a *argList) addString(name, value string) {
	if value == "" {
		return
	}
	a.args = append(a.args, name, value)
}

func (a *argList) addBool(name string, value bool) {
	if !value {
		return
	}
	a.args = append(a.args, name)
}

func (a *argList) addStrings(name string, values []string) {
	for _, value := range values {
		a.addString(name, value)
	}
}

func (a *argList) addArgs(values []string) {
	a.args = append(a.args, values...)
}
