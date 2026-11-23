package formula

// ShiftRefs returns the formula text with its relative references moved by
// dCol columns and dRow rows. It is the adjustment that a copy, a fill or a
// sort makes to a formula that moves by that distance.
//
// Absolute parts ($A, $1) stay. A column-only reference shifts only on the
// column axis and a row-only reference only on the row axis. A reference
// with a part that leaves the sheet becomes #REF!. All other text is
// unchanged. src is the formula without the leading '='.
func ShiftRefs(src string, dCol, dRow int) string {
	if dCol == 0 && dRow == 0 {
		return src
	}
	return RewriteRefs(src, func(site RefSite) (RefSite, RefAction) {
		from, fromMoved, ok := shiftRefPoint(site.From, dCol, dRow)
		if !ok {
			return site, RefInvalid
		}
		to, toMoved := from, fromMoved
		if site.Kind != RefCell {
			to, toMoved, ok = shiftRefPoint(site.To, dCol, dRow)
			if !ok {
				return site, RefInvalid
			}
		}
		if !fromMoved && !toMoved {
			return site, RefKeep
		}
		site.From, site.To = from, to
		return site, RefReplace
	})
}

// shiftRefPoint moves the relative parts of p. ok is false when a part
// leaves the sheet.
func shiftRefPoint(p RefPoint, dCol, dRow int) (out RefPoint, moved, ok bool) {
	if p.Col > 0 && !p.AbsCol && dCol != 0 {
		p.Col += dCol
		if p.Col < 1 || p.Col > refMaxCol {
			return RefPoint{}, false, false
		}
		moved = true
	}
	if p.Row > 0 && !p.AbsRow && dRow != 0 {
		p.Row += dRow
		if p.Row < 1 || p.Row > refMaxRow {
			return RefPoint{}, false, false
		}
		moved = true
	}
	return p, moved, true
}
