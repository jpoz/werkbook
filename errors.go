package werkbook

import "errors"

var (
	// ErrInvalidCellRef is returned when a cell reference string is invalid.
	ErrInvalidCellRef = errors.New("invalid cell reference")
	// ErrSheetNotFound is returned when a referenced sheet does not exist.
	ErrSheetNotFound = errors.New("sheet not found")
	// ErrUnsupportedType is returned when a value type cannot be stored in a cell.
	ErrUnsupportedType = errors.New("unsupported value type")
)
