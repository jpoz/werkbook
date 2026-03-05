package main

func cmdCapabilities(args []string, globals globalFlags) int {
	if hasHelpFlag(args) {
		return writeHelpTopic([]string{"capabilities"}, globals)
	}
	if len(args) > 0 {
		writeError("capabilities", errUsage("capabilities does not accept positional arguments"), globals)
		return ExitUsage
	}

	writeSuccess("capabilities", wbToolSpec(), globals)
	return ExitSuccess
}
