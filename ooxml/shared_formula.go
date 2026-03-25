package ooxml

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/formula"
)

// sharedFormulaMaster holds the master cell info for a shared formula group.
type sharedFormulaMaster struct {
	formula string
	col     int
	row     int
}

// expandSharedFormulas resolves shared formula references within a SheetData.
// Excel optimises repeated formulas by storing the text only on a "master" cell
// (the one with both t="shared" and a ref= attribute) and marking every other
// cell in the group with just the si= index. This function finds each master,
// then fills in the formula text on every child cell by shifting cell references
// according to the child's offset from the master.
//
// After expansion every formerly-shared cell becomes a normal (standalone)
// formula cell, so the rest of the pipeline does not need to know about shared
// formulas at all.
func expandSharedFormulas(sd *SheetData) {
	// First pass: collect masters keyed by si index.
	masters := make(map[int]sharedFormulaMaster)
	for _, row := range sd.Rows {
		for _, cell := range row.Cells {
			if cell.SharedIndex >= 0 && cell.Formula != "" {
				col, r, err := cellRefToCoordinates(cell.Ref)
				if err != nil {
					continue
				}
				masters[cell.SharedIndex] = sharedFormulaMaster{
					formula: cell.Formula,
					col:     col,
					row:     r,
				}
			}
		}
	}

	if len(masters) == 0 {
		return
	}

	// Second pass: expand children and convert masters to standalone.
	for i := range sd.Rows {
		for j := range sd.Rows[i].Cells {
			cell := &sd.Rows[i].Cells[j]
			if cell.SharedIndex < 0 {
				continue
			}
			if cell.Formula != "" {
				// Master cell — convert to standalone.
				cell.FormulaType = ""
				cell.FormulaRef = ""
				cell.SharedIndex = -1
				continue
			}
			// Child cell — look up master and shift.
			master, ok := masters[cell.SharedIndex]
			if !ok {
				continue
			}
			col, r, err := cellRefToCoordinates(cell.Ref)
			if err != nil {
				continue
			}
			dCol := col - master.col
			dRow := r - master.row
			cell.Formula = shiftFormulaRefs(master.formula, dCol, dRow)
			cell.FormulaType = ""
			cell.FormulaRef = ""
			cell.SharedIndex = -1
		}
	}
}

// shiftFormulaRefs tokenizes a formula string and shifts all non-absolute cell
// references by (dCol, dRow). It preserves absolute references marked with $.
func shiftFormulaRefs(src string, dCol, dRow int) string {
	tokens, err := formula.Tokenize(src)
	if err != nil {
		return src // if we can't parse, return as-is
	}

	var b strings.Builder
	b.Grow(len(src) + 16)
	lastPos := 0

	for _, tok := range tokens {
		if tok.Type != formula.TokCellRef {
			continue
		}
		ref, err := parseCellRefForShift(tok.Value)
		if err != nil {
			continue
		}

		newCol := ref.col
		newRow := ref.row
		if !ref.absCol {
			newCol += dCol
		}
		if !ref.absRow && ref.row > 0 {
			newRow += dRow
		}
		if newCol < 1 || newRow < 0 {
			continue // out of range, skip shifting
		}

		shifted := buildCellRefString(ref, newCol, newRow)

		// Replace the token in the source string.
		b.WriteString(src[lastPos:tok.Pos])
		b.WriteString(shifted)
		lastPos = tok.Pos + len(tok.Value)
	}

	if lastPos == 0 {
		return src // nothing shifted
	}
	b.WriteString(src[lastPos:])
	return b.String()
}

// cellRefParts holds parsed components of a cell reference token for shifting.
type cellRefParts struct {
	prefix string // sheet prefix including "!" (e.g. "Sheet1!" or "'My Sheet'!")
	absCol bool
	col    int
	absRow bool
	row    int // 0 for column-only refs like F:F
}

// parseCellRefForShift parses a cell reference token value into its components.
func parseCellRefForShift(raw string) (cellRefParts, error) {
	var p cellRefParts
	s := raw

	// Extract sheet prefix (everything up to and including '!').
	bangIdx := -1
	if len(s) > 0 && s[0] == '\'' {
		// Quoted sheet reference.
		i := 1
		for i < len(s) {
			if s[i] == '\'' {
				if i+1 < len(s) && s[i+1] == '\'' {
					i += 2
					continue
				}
				break
			}
			i++
		}
		if i+1 < len(s) && s[i+1] == '!' {
			bangIdx = i + 1
		}
	} else {
		bangIdx = strings.IndexByte(s, '!')
	}

	if bangIdx >= 0 {
		p.prefix = s[:bangIdx+1]
		s = s[bangIdx+1:]
	}

	// Parse optional $, column letters, optional $, row digits.
	i := 0
	if i < len(s) && s[i] == '$' {
		p.absCol = true
		i++
	}

	colStart := i
	for i < len(s) && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == colStart {
		return p, fmt.Errorf("no column in %q", raw)
	}
	p.col = formula.ColLettersToNumber(s[colStart:i])

	if i < len(s) && s[i] == '$' {
		p.absRow = true
		i++
	}

	if i < len(s) && s[i] >= '0' && s[i] <= '9' {
		row := 0
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			row = row*10 + int(s[i]-'0')
			i++
		}
		p.row = row
	}

	return p, nil
}

// buildCellRefString reconstructs a cell reference string from its parts.
func buildCellRefString(p cellRefParts, col, row int) string {
	var b strings.Builder
	b.WriteString(p.prefix)
	if p.absCol {
		b.WriteByte('$')
	}
	b.WriteString(formula.ColNumberToLetters(col))
	if p.absRow {
		b.WriteByte('$')
	}
	if row > 0 {
		b.WriteString(strconv.Itoa(row))
	}
	return b.String()
}

// cellRefToCoordinates extracts (col, row) from a cell reference string like "F7".
func cellRefToCoordinates(ref string) (int, int, error) {
	i := 0
	for i < len(ref) && ((ref[i] >= 'A' && ref[i] <= 'Z') || (ref[i] >= 'a' && ref[i] <= 'z')) {
		i++
	}
	if i == 0 || i == len(ref) {
		return 0, 0, fmt.Errorf("invalid cell ref %q", ref)
	}
	col := formula.ColLettersToNumber(ref[:i])
	row := 0
	for j := i; j < len(ref); j++ {
		if ref[j] < '0' || ref[j] > '9' {
			return 0, 0, fmt.Errorf("invalid cell ref %q", ref)
		}
		row = row*10 + int(ref[j]-'0')
	}
	if row == 0 {
		return 0, 0, fmt.Errorf("invalid cell ref %q: row is 0", ref)
	}
	return col, row, nil
}
