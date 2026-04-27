package formula

import (
	"math"
	"strings"
)

var offsetScalarArgSpec = ArgSpec{Adapt: ArgAdaptScalarizeAny}
var offsetRefOnlyArgSpec = ArgSpec{Adapt: ArgAdaptLegacyIntersectRef}
var indirectRefProducerSpec = refProducerFuncSpec(evalINDIRECT)
var offsetRefProducerSpec = refProducerFuncSpec(evalOFFSET)

func refProducerFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindLookup,
		Args: []ArgSpec{{
			Load:  ArgLoadPassthrough,
			Adapt: ArgAdaptPassThrough,
		}},
		VarArg: func(_ int) ArgSpec {
			return ArgSpec{
				Load:  ArgLoadPassthrough,
				Adapt: ArgAdaptPassThrough,
			}
		},
		Return: ReturnModeRef,
		Eval:   eval,
	}
}

func callRefProducerWithSpec(spec FuncSpec, args []Value, ctx *EvalContext) (Value, error) {
	got, err := spec.Eval(loadEvalFuncArgs(spec, args, ctx), ctx)
	if err != nil {
		return Value{}, err
	}
	return EvalValueToValue(got), nil
}

func evalINDIRECT(args []EvalValue, ctx *EvalContext) (EvalValue, error) {
	if len(args) < 1 || len(args) > 2 {
		return evalError(ErrValVALUE), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}
	if ctx == nil || ctx.Resolver == nil {
		return evalError(ErrValREF), nil
	}

	a1Style := true
	if len(args) == 2 {
		a1Style = IsTruthy(EvalValueToValue(args[1]))
	}
	legacyArg := EvalValueToValue(args[0])
	if legacyArg.Type == ValueArray {
		return evalIndirectLegacyArray(legacyArg, a1Style, ctx), nil
	}
	refText := ValueToString(legacyArg)
	if refText == "" {
		return evalError(ErrValREF), nil
	}
	return evalIndirectText(refText, a1Style, ctx), nil
}

func evalIndirectLegacyArray(arg Value, a1Style bool, ctx *EvalContext) EvalValue {
	rows := make([][]Value, len(arg.Array))
	for r, srcRow := range arg.Array {
		row := make([]Value, len(srcRow))
		for c, cell := range srcRow {
			got := evalIndirectText(ValueToString(cell), a1Style, ctx)
			row[c] = indirectElementValue(got, ctx)
		}
		rows[r] = row
	}
	return evalArray(rows, SpillBounded)
}

func indirectElementValue(v EvalValue, ctx *EvalContext) Value {
	if v.Kind == EvalRef && v.Ref != nil {
		if v.Ref.FromCol == v.Ref.ToCol && v.Ref.FromRow == v.Ref.ToRow {
			if v.Ref.Legacy != nil && v.Ref.Legacy.SingleCellValue.Type != ValueEmpty {
				out := v.Ref.Legacy.SingleCellValue
				out.evalRef = cloneRefValue(v.Ref)
				return out
			}
			if v.Ref.Materialized != nil {
				out := EvalValueToValue(v.Ref.Materialized.Cell(0, 0))
				out.evalRef = cloneRefValue(v.Ref)
				return out
			}
		}
	}
	legacy := EvalValueToValue(v)
	if legacy.Type == ValueArray {
		return explicitIntersect(legacy, ctx)
	}
	return legacy
}

func evalIndirectText(refText string, a1Style bool, ctx *EvalContext) EvalValue {
	if refText == "" {
		return evalError(ErrValREF)
	}
	if !a1Style {
		converted, err := r1c1ToA1At(refText, ctx.CurrentRow, ctx.CurrentCol)
		if err != nil {
			return evalError(ErrValREF)
		}
		refText = converted
	}

	cleaned := strings.ReplaceAll(refText, "$", "")
	sheet := ""
	prefix, cellPart := splitSheetPrefix(cleaned)
	if prefix != "" {
		sheetPart := prefix[:len(prefix)-1]
		if len(sheetPart) >= 2 && sheetPart[0] == '\'' && sheetPart[len(sheetPart)-1] == '\'' {
			sheetPart = strings.ReplaceAll(sheetPart[1:len(sheetPart)-1], "''", "'")
		}
		sheet = sheetPart
	}

	if colonIdx := strings.IndexByte(cellPart, ':'); colonIdx >= 0 {
		left := cellPart[:colonIdx]
		right := cellPart[colonIdx+1:]
		addr, err := indirectParseRange(left, right, sheet)
		if err != nil {
			return evalError(ErrValREF)
		}
		isFullCol := addr.FromRow == 1 && addr.ToRow >= maxRows
		isFullRow := addr.FromCol == 1 && addr.ToCol >= maxCols
		checkRows := addr.ToRow - addr.FromRow + 1
		checkCols := addr.ToCol - addr.FromCol + 1
		var legacy *RefLegacyBoundary
		if isFullRow {
			checkCols = 1
			legacy = &RefLegacyBoundary{
				PlaceholderRows: checkRows,
				PlaceholderCols: 1,
				UseEmptyArray:   true,
			}
		} else if isFullCol {
			checkRows = 1
			legacy = &RefLegacyBoundary{
				PlaceholderRows: 1,
				PlaceholderCols: checkCols,
				UseEmptyArray:   true,
			}
		}
		if RangeCellCountExceedsLimit(checkRows, checkCols) {
			return evalError(ErrValREF)
		}

		var rows [][]Value
		if !isFullCol && !isFullRow {
			rows = ctx.Resolver.GetRangeValues(addr)
			if isRangeOverflowMatrix(rows) {
				return evalError(ErrValREF)
			}
		}
		return newEvalRangeRef(addr, rows, ctx.Resolver, legacy)
	}

	col, row, err := indirectParseCell(cellPart)
	if err == nil {
		addr := CellAddr{Sheet: sheet, Col: col, Row: row}
		val := ctx.Resolver.GetCellValue(addr)
		return newEvalSingleCellRef(addr, val)
	}
	if nameResolver, ok := ctx.Resolver.(DefinedNameResolver); ok {
		scopeSheet := ctx.CurrentSheet
		if sheet != "" {
			scopeSheet = sheet
		}
		if named, ok := nameResolver.ResolveDefinedNameValue(cellPart, scopeSheet); ok {
			return valueToIndirectEvalValue(named, ctx.Resolver)
		}
	}
	return evalError(ErrValREF)
}

func evalOFFSET(args []EvalValue, ctx *EvalContext) (EvalValue, error) {
	if len(args) < 3 || len(args) > 5 {
		return evalError(ErrValVALUE), nil
	}
	if ctx == nil || ctx.Resolver == nil {
		return evalError(ErrValREF), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}
	if args[0].Kind != EvalRef || args[0].Ref == nil {
		return evalError(ErrValVALUE), nil
	}

	ref := args[0].Ref
	sheet := ref.Sheet
	fromRow := ref.FromRow
	fromCol := ref.FromCol
	toRow := ref.ToRow
	toCol := ref.ToCol

	refHeight := toRow - fromRow + 1
	refWidth := toCol - fromCol + 1
	height := refHeight
	width := refWidth
	candidateHeight := refHeight
	candidateWidth := refWidth

	if len(args) >= 4 {
		heightArg := offsetLegacyArg(args[3], offsetScalarArgSpec, ctx)
		if heightArg.Type != ValueEmpty {
			hN, errV := CoerceNum(heightArg)
			if errV != nil {
				return evalScalar(*errV), nil
			}
			candidateHeight = int(math.Trunc(hN))
		}
	}
	if len(args) >= 5 {
		widthArg := offsetLegacyArg(args[4], offsetScalarArgSpec, ctx)
		if widthArg.Type != ValueEmpty {
			wN, errV := CoerceNum(widthArg)
			if errV != nil {
				return evalScalar(*errV), nil
			}
			candidateWidth = int(math.Trunc(wN))
		}
	}
	sizeSpec := offsetRefOnlyArgSpec
	rowsSpec := offsetRefOnlyArgSpec
	colsSpec := offsetRefOnlyArgSpec
	if candidateHeight == 1 && candidateWidth == 1 {
		sizeSpec = offsetScalarArgSpec
		rowsSpec = offsetScalarArgSpec
		colsSpec = offsetScalarArgSpec
	}

	if len(args) >= 4 {
		heightArg := offsetLegacyArg(args[3], sizeSpec, ctx)
		if heightArg.Type != ValueEmpty {
			hN, errV := CoerceNum(heightArg)
			if errV != nil {
				return evalScalar(*errV), nil
			}
			height = int(math.Trunc(hN))
		}
	}
	if len(args) >= 5 {
		widthArg := offsetLegacyArg(args[4], sizeSpec, ctx)
		if widthArg.Type != ValueEmpty {
			wN, errV := CoerceNum(widthArg)
			if errV != nil {
				return evalScalar(*errV), nil
			}
			width = int(math.Trunc(wN))
		}
	}

	// Legacy OFFSET scalarizes anonymous row/col arrays only when the result
	// remains a single cell; multi-cell refs must reject them instead.
	rowsN, errV := CoerceNum(offsetLegacyArg(args[1], rowsSpec, ctx))
	if errV != nil {
		return evalScalar(*errV), nil
	}
	rowsOff := int(math.Trunc(rowsN))

	colsN, errV := CoerceNum(offsetLegacyArg(args[2], colsSpec, ctx))
	if errV != nil {
		return evalScalar(*errV), nil
	}
	colsOff := int(math.Trunc(colsN))
	if height == 0 || width == 0 {
		return evalError(ErrValREF), nil
	}

	newFromRow := fromRow + rowsOff
	newFromCol := fromCol + colsOff
	newToRow := newFromRow
	newToCol := newFromCol
	if height > 0 {
		newToRow = newFromRow + height - 1
	} else {
		newToRow = newFromRow
		newFromRow = newFromRow + height + 1
		height = -height
	}
	if width > 0 {
		newToCol = newFromCol + width - 1
	} else {
		newToCol = newFromCol
		newFromCol = newFromCol + width + 1
		width = -width
	}

	if newFromRow < 1 || newFromCol < 1 || newToRow > maxRows || newToCol > maxCols {
		return evalError(ErrValREF), nil
	}

	if height == 1 && width == 1 {
		addr := CellAddr{Sheet: sheet, Col: newFromCol, Row: newFromRow}
		val := ctx.Resolver.GetCellValue(addr)
		val.FromCell = true
		return newEvalSingleCellRef(addr, val), nil
	}

	addr := RangeAddr{
		Sheet:   sheet,
		FromCol: newFromCol,
		FromRow: newFromRow,
		ToCol:   newToCol,
		ToRow:   newToRow,
	}
	rows := ctx.Resolver.GetRangeValues(addr)
	if isRangeOverflowMatrix(rows) {
		return evalError(ErrValREF), nil
	}
	return newEvalRangeRef(addr, rows, ctx.Resolver, nil), nil
}

func offsetLegacyArg(arg EvalValue, spec ArgSpec, ctx *EvalContext) Value {
	return adaptFuncArg(spec, EvalValueToValue(arg), ctx)
}

func valueToIndirectEvalValue(v Value, resolver CellResolver) EvalValue {
	ev := valueToEvalValueWithResolver(v, resolver)
	if ev.Kind != EvalRef || ev.Ref == nil || ev.Ref.Legacy != nil {
		return ev
	}
	rows := ev.Ref.ToRow - ev.Ref.FromRow + 1
	cols := ev.Ref.ToCol - ev.Ref.FromCol + 1
	switch {
	case ev.Ref.FromCol == 1 && ev.Ref.ToCol >= maxCols:
		ev.Ref.Legacy = &RefLegacyBoundary{
			PlaceholderRows: rows,
			PlaceholderCols: 1,
			UseEmptyArray:   true,
		}
	case ev.Ref.FromRow == 1 && ev.Ref.ToRow >= maxRows:
		ev.Ref.Legacy = &RefLegacyBoundary{
			PlaceholderRows: 1,
			PlaceholderCols: cols,
			UseEmptyArray:   true,
		}
	}
	return ev
}

func newEvalRangeRef(addr RangeAddr, rows [][]Value, resolver CellResolver, legacy *RefLegacyBoundary) EvalValue {
	return EvalValue{
		Kind: EvalRef,
		Ref: &RefValue{
			Sheet:        addr.Sheet,
			SheetEnd:     addr.SheetEnd,
			FromCol:      addr.FromCol,
			FromRow:      addr.FromRow,
			ToCol:        addr.ToCol,
			ToRow:        addr.ToRow,
			Materialized: newResolverRangeGrid(addr, rows, resolver),
			Legacy:       legacy,
		},
	}
}

func newEvalSingleCellRef(addr CellAddr, value Value) EvalValue {
	value.CellOrigin = ptrCell(addr)
	return EvalValue{
		Kind: EvalRef,
		Ref: &RefValue{
			Sheet:        addr.Sheet,
			SheetEnd:     addr.SheetEnd,
			FromCol:      addr.Col,
			FromRow:      addr.Row,
			ToCol:        addr.Col,
			ToRow:        addr.Row,
			Materialized: newLegacyValueGrid([][]Value{{stripRefMetadata(value)}}),
			Legacy: &RefLegacyBoundary{
				SingleCellValue: value,
			},
		},
	}
}
