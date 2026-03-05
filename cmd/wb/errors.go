package main

import "fmt"

// Exit codes.
const (
	ExitSuccess  = 0
	ExitFileIO   = 1
	ExitValidate = 2
	ExitPartial  = 3
	ExitUsage    = 4
	ExitInternal = 99
)

// Error code constants.
const (
	ErrCodeFileNotFound   = "FILE_NOT_FOUND"
	ErrCodeFileOpenFailed = "FILE_OPEN_FAILED"
	ErrCodeFileSaveFailed = "FILE_SAVE_FAILED"
	ErrCodeSheetNotFound  = "SHEET_NOT_FOUND"
	ErrCodeInvalidRange   = "INVALID_RANGE"
	ErrCodeInvalidPatch   = "INVALID_PATCH"
	ErrCodeInvalidSpec    = "INVALID_SPEC"
	ErrCodeInvalidFormat  = "INVALID_FORMAT"
	ErrCodeUsage          = "USAGE"
	ErrCodeInternal       = "INTERNAL"
	ErrCodePartialFailure = "PARTIAL_FAILURE"
	ErrCodeValidation     = "VALIDATION"
	ErrCodeEncryptedFile  = "ENCRYPTED_FILE"
)

func errFileNotFound(path string, err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeFileNotFound,
		Message: fmt.Sprintf("could not open %q: %v", path, err),
		Hint:    "Check the file path. Use 'wb create' to create a new file.",
	}
}

func errFileOpen(path string, err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeFileOpenFailed,
		Message: fmt.Sprintf("could not open %q: %v", path, err),
		Hint:    "Ensure the file is a valid .xlsx file.",
	}
}

func errFileSave(path string, err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeFileSaveFailed,
		Message: fmt.Sprintf("could not save %q: %v", path, err),
		Hint:    "Check file permissions and disk space.",
	}
}

func errSheetNotFound(name string) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeSheetNotFound,
		Message: fmt.Sprintf("sheet %q not found", name),
		Hint:    "Use 'wb info' to list available sheets.",
	}
}

func errInvalidRange(ref string, err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeInvalidRange,
		Message: fmt.Sprintf("invalid range %q: %v", ref, err),
		Hint:    "Use A1 notation, e.g. 'A1:D10'.",
	}
}

func errInvalidPatch(err error) *ErrorInfo {
	msg := "no patch data provided"
	if err != nil {
		msg = fmt.Sprintf("invalid patch JSON: %v", err)
	}
	return &ErrorInfo{
		Code:    ErrCodeInvalidPatch,
		Message: msg,
		Hint:    "Provide a JSON array of patch operations via --patch or stdin.",
	}
}

func errInvalidSpec(err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeInvalidSpec,
		Message: fmt.Sprintf("invalid spec JSON: %v", err),
		Hint:    "Provide a JSON object with 'sheets', 'cells', and/or 'rows' fields. Run 'wb create --help' for the spec format.",
	}
}

func errUsage(msg string) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeUsage,
		Message: msg,
		Hint:    "Run 'wb <command> --help' for command usage.",
	}
}

func errEncryptedFile(path string) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeEncryptedFile,
		Message: fmt.Sprintf("file %q is password-protected or encrypted", path),
		Hint:    "Remove the password in Excel (File > Info > Protect Workbook > Encrypt with Password) and try again.",
	}
}

func errValidation(msg string) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeValidation,
		Message: msg,
	}
}

func errInternal(err error) *ErrorInfo {
	return &ErrorInfo{
		Code:    ErrCodeInternal,
		Message: fmt.Sprintf("internal error: %v", err),
	}
}
