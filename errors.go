package werkbook

import (
	"errors"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

var (
	// ErrInvalidCellRef is returned when a cell reference string is invalid.
	ErrInvalidCellRef = errors.New("invalid cell reference")
	// ErrSheetNotFound is returned when a referenced sheet does not exist.
	ErrSheetNotFound = errors.New("sheet not found")
	// ErrUnsupportedType is returned when a value type cannot be stored in a cell.
	ErrUnsupportedType = errors.New("unsupported value type")
	// ErrFormulaTooLarge is returned when a formula or formula expansion exceeds
	// the parser's size budget.
	ErrFormulaTooLarge = formula.ErrFormulaTooLarge
	// ErrEncryptedFile is returned when the file is encrypted/password-protected.
	ErrEncryptedFile = ooxml.ErrEncryptedFile
)
