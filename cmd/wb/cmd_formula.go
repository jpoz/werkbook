package main

import (
	"sort"

	"github.com/jpoz/werkbook/formula"
)

type formulaListData struct {
	Functions []formula.FunctionInfo `json:"functions"`
	Count     int                    `json:"count"`
}

func cmdFormula(args []string, globals globalFlags) int {
	cmd := "formula"

	if hasHelpFlag(args) {
		if len(args) > 0 && args[0] == "list" {
			return writeHelpTopic([]string{cmd, "list"}, globals)
		}
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON) {
		return ExitUsage
	}

	if len(args) == 0 {
		writeError(cmd, errUsage("subcommand required: 'formula list'"), globals)
		return ExitUsage
	}

	switch args[0] {
	case "list":
		return cmdFormulaList(globals)
	default:
		writeError(cmd, errUsage("unknown subcommand: "+args[0]+". Available: list"), globals)
		return ExitUsage
	}
}

func cmdFormulaList(globals globalFlags) int {
	infos := formula.RegisteredFunctionInfos()
	sort.Slice(infos, func(i, j int) bool { return infos[i].Name < infos[j].Name })

	data := formulaListData{
		Functions: infos,
		Count:     len(infos),
	}
	writeSuccess("formula", data, globals)
	return ExitSuccess
}
