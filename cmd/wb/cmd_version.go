package main

var version = "dev"

type versionData struct {
	Version string `json:"version"`
}

func cmdVersion(args []string, globals globalFlags) int {
	if hasHelpFlag(args) {
		return writeHelpTopic([]string{"version"}, globals)
	}
	if len(args) > 0 {
		writeError("version", errUsage("version does not accept positional arguments"), globals)
		return ExitUsage
	}

	data := versionData{Version: version}
	writeSuccess("version", data, globals)
	return ExitSuccess
}
