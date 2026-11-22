package werkbook

import (
	"sort"

	"github.com/jpoz/werkbook/formula"
)

// maxEvalRecursionDepth is the number of nested formula evaluations that the
// engine permits on the Go stack. Evaluation is recursive: a formula reads a
// precedent cell, and a stale precedent is evaluated in a nested call. Each
// level uses tens of kilobytes of stack, so a long dependency chain (for
// example a running total where each row reads the row above) can exceed the
// 1 GB goroutine stack limit. That failure is a fatal error that a caller
// cannot recover from.
//
// At this depth the engine stops the recursion and evaluates the remaining
// cells with an explicit stack. See evalDeferred.
const maxEvalRecursionDepth = 256

// deferredStackCompactMin is the smallest number of dead entries for which
// evalDeferred compacts its stack.
const deferredStackCompactMin = 1024

// deferredEval is one cell on the explicit evaluation stack of evalDeferred.
type deferredEval struct {
	sheet    *Sheet
	cell     *Cell
	col, row int
	// raw is true when a reader needs the raw (array) result of the cell,
	// not only its display value.
	raw bool
	// waiting is true after an attempt to evaluate the cell found stale
	// precedents. Those precedents are above the cell on the stack, and the
	// cell is in File.evaluating, so a read of it is a circular reference.
	waiting bool
	// dead is true when a newer entry for the same cell is higher on the
	// stack.
	dead bool
}

func (d *deferredEval) key() cellKey {
	return cellKey{sheet: d.sheet.name, col: d.col, row: d.row}
}

// stale reports whether a reader of this entry would start an evaluation.
func (d *deferredEval) stale(calcGen uint64) bool {
	if d.cell.needsEval(calcGen) {
		return true
	}
	return d.raw && d.cell.rawStale(calcGen)
}

// evalProbe is the state of one evaluation attempt in evalDeferred. During an
// attempt no nested evaluation starts. A read of a stale cell is recorded as
// a miss, and the reader gets the old value of the cell. The result of an
// attempt with misses is discarded.
type evalProbe struct {
	// read counts the value reads that the formula engine requested from
	// the resolver in this attempt.
	read int
	// missedRead is the read that recorded the misses.
	missedRead int
	missed     []deferredEval
	missedIdx  map[cellKey]int
}

func (p *evalProbe) reset() {
	p.read = 0
	p.missedRead = 0
	p.missed = p.missed[:0]
	clear(p.missedIdx)
}

// needsEval reports whether the cell has a formula whose cached value is not
// current for the given calculation generation.
func (c *Cell) needsEval(calcGen uint64) bool {
	return c.formula != "" && (c.dirty || c.cachedGen < calcGen)
}

// rawStale reports whether the cached raw result of a formula cell is not
// current for the given calculation generation.
func (c *Cell) rawStale(calcGen uint64) bool {
	return c.formula != "" && (c.dirty || c.rawCachedGen != calcGen)
}

// beginResolverRead marks the start of one value read by the formula engine.
func (f *File) beginResolverRead() {
	if f.probe != nil {
		f.probe.read++
	}
}

// deferEval is called before a nested evaluation of a stale cell starts. It
// returns false when the evaluation must go on as usual. It returns true
// during an attempt of evalDeferred: the caller must not evaluate the cell
// and must use the old value, because the result of the attempt is discarded.
//
// Only the first read that finds stale cells records its misses. Reads after
// it use old values, so the cells that they request are not reliable: a
// branch that the formula does not take with the correct values must not
// cause an evaluation.
func (f *File) deferEval(s *Sheet, c *Cell, col, row int, raw bool) bool {
	p := f.probe
	if p == nil {
		return false
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	if f.evaluating[key] {
		// A circular reference. The usual path reports it.
		return false
	}
	if len(p.missed) > 0 && p.read != p.missedRead {
		return true
	}
	p.missedRead = p.read
	if idx, ok := p.missedIdx[key]; ok {
		p.missed[idx].raw = p.missed[idx].raw || raw
		return true
	}
	if p.missedIdx == nil {
		p.missedIdx = make(map[cellKey]int)
	}
	p.missedIdx[key] = len(p.missed)
	p.missed = append(p.missed, deferredEval{sheet: s, cell: c, col: col, row: row, raw: raw})
	return true
}

// evalCellOutcomeBounded evaluates a formula cell. When the recursion is
// deeper than maxEvalRecursionDepth, it uses evalDeferred. The caller must
// have the cell in f.evaluating.
func (s *Sheet) evalCellOutcomeBounded(c *Cell, col, row int) (CellEvalOutcome, error) {
	f := s.file
	if f.probe == nil && len(f.evaluating) > maxEvalRecursionDepth {
		return f.evalDeferred(s, c, col, row)
	}
	return s.evalCellOutcome(c, col, row)
}

// evalDeferred evaluates a formula cell and the stale cells that it reads,
// directly or through other cells, without recursion. The length of the
// dependency chain does not increase the depth of the Go stack.
//
// It attempts to evaluate the cell on the top of an explicit stack. If the
// attempt reads stale cells, its result is discarded, the stale cells go on
// the stack, and the cell is attempted again after they are resolved. The
// stack is driven by the reads of the evaluator, so a cell is evaluated only
// if a formula reads its value. That is the same set of cells that the
// recursive path evaluates. A reference that is used only for its address
// (COLUMN(B1)) or that is in a branch that is not taken does not cause an
// evaluation.
//
// A cell that waits for its precedents stays in f.evaluating. A read of such
// a cell is a circular reference, and it gives the same error value as in
// the recursive path.
func (f *File) evalDeferred(s *Sheet, c *Cell, col, row int) (CellEvalOutcome, error) {
	probe := &evalProbe{}
	f.probe = probe
	// The root is in f.evaluating already; its caller removes it.
	stack := []deferredEval{{sheet: s, cell: c, col: col, row: row, waiting: true}}
	defer func() {
		f.probe = nil
		for i := 1; i < len(stack); i++ {
			if stack[i].waiting {
				delete(f.evaluating, stack[i].key())
			}
		}
	}()

	// live has the stack index of the entry to evaluate for each cell that
	// is on the stack and not attempted.
	live := make(map[cellKey]int)
	dead := 0

	for {
		top := &stack[len(stack)-1]
		isRoot := len(stack) == 1
		key := top.key()
		if !isRoot {
			if top.dead {
				stack = stack[:len(stack)-1]
				dead--
				continue
			}
			if !top.waiting {
				delete(live, key)
				if !top.stale(f.calcGen) {
					// Resolved through another entry.
					stack = stack[:len(stack)-1]
					continue
				}
				top.waiting = true
				f.evaluating[key] = true
			}
		}

		probe.reset()
		outcome, err := top.sheet.evalCellOutcome(top.cell, top.col, top.row)
		if len(probe.missed) == 0 {
			if isRoot {
				return outcome, err
			}
			top.sheet.storeDeferredResult(top.cell, top.col, top.row, outcome, err)
			delete(f.evaluating, key)
			stack = stack[:len(stack)-1]
			continue
		}

		// Most formulas read cells above them or to the left. Put those
		// cells on top, so they are resolved first and the cells below them
		// on the stack find their precedents ready.
		missed := probe.missed
		sort.Slice(missed, func(i, j int) bool {
			a, b := &missed[i], &missed[j]
			if a.sheet != b.sheet {
				return a.sheet.name > b.sheet.name
			}
			if a.row != b.row {
				return a.row > b.row
			}
			return a.col > b.col
		})
		for _, m := range missed {
			mk := m.key()
			if idx, ok := live[mk]; ok {
				// The cell is on the stack below a cell that needs it
				// first. Replace the old entry.
				stack[idx].dead = true
				dead++
				m.raw = m.raw || stack[idx].raw
			}
			live[mk] = len(stack)
			stack = append(stack, m)
		}

		if dead >= deferredStackCompactMin && dead > len(stack)/2 {
			kept := stack[:0]
			for _, e := range stack {
				if e.dead {
					continue
				}
				if !e.waiting {
					live[e.key()] = len(kept)
				}
				kept = append(kept, e)
			}
			stack = kept
			dead = 0
		}
	}
}

// storeDeferredResult caches the result of a cell that evalDeferred
// evaluated. It gives the cell the same state as a nested evaluation gives.
func (s *Sheet) storeDeferredResult(c *Cell, col, row int, outcome CellEvalOutcome, err error) {
	if err != nil {
		c.value = evalErrorValue(c, err)
		c.cachedGen = s.file.calcGen
		c.dirty = false
		// evaluateFormulaRaw gives this value for a failed evaluation. It
		// is cached so that a reader of the raw result does not defer the
		// cell again.
		c.rawValue = formula.ErrorVal(formula.ErrValVALUE)
		c.rawCachedGen = s.file.calcGen
		return
	}
	s.applyCellOutcome(c, col, row, outcome)
}

// evalErrorValue classifies an evaluation failure (parse or compile failure
// gives #NAME?; expansion overflow or a runtime engine failure gives #VALUE!).
// The underlying message stays reachable through errors.Is and errors.As on
// an *formula.EvalError.
func evalErrorValue(c *Cell, err error) Value {
	code := formula.ErrorValueFromErr(err).String()
	if c.value.Type == TypeString {
		return Value{Type: TypeString, String: code}
	}
	return Value{Type: TypeError, String: code}
}
