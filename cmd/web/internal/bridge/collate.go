package bridge

import (
	"strings"
	"unicode"
)

// Text order for a sort. A compare of bytes puts each letter with an accent
// after "z" ("Émile" after "Zeta"), and puts punctuation before digits. A
// spreadsheet sorts text as a reader expects: digits first, then symbols,
// then letters, where a letter with an accent sorts with its base letter and
// letter case is not important.
//
// The standard library has no collation, so this file has a small one: a
// weight for each character class and a table of base letters for the Latin
// letters of Western and Central European languages (Latin-1 and Latin
// Extended-A). Letters of other scripts (Greek, Cyrillic, ...) sort by their
// code point, after the Latin letters. It is not a full Unicode collation:
// the rule of Excel that a hyphen and an apostrophe are ignored is not
// applied.

// Character classes, in sort order.
const (
	collateDigit = iota
	collateSymbol
	collateLetter
)

// latinBase maps a lower-case Latin letter with a mark, or a ligature, to
// its base letters.
var latinBase = map[rune]string{
	// Latin-1 Supplement.
	'à': "a", 'á': "a", 'â': "a", 'ã': "a", 'ä': "a", 'å': "a", 'æ': "ae",
	'ç': "c",
	'è': "e", 'é': "e", 'ê': "e", 'ë': "e",
	'ì': "i", 'í': "i", 'î': "i", 'ï': "i",
	'ð': "d", 'ñ': "n",
	'ò': "o", 'ó': "o", 'ô': "o", 'õ': "o", 'ö': "o", 'ø': "o",
	'ù': "u", 'ú': "u", 'û': "u", 'ü': "u",
	'ý': "y", 'ÿ': "y", 'þ': "th", 'ß': "ss",
	// Latin Extended-A.
	'ā': "a", 'ă': "a", 'ą': "a",
	'ć': "c", 'ĉ': "c", 'ċ': "c", 'č': "c",
	'ď': "d", 'đ': "d",
	'ē': "e", 'ĕ': "e", 'ė': "e", 'ę': "e", 'ě': "e",
	'ĝ': "g", 'ğ': "g", 'ġ': "g", 'ģ': "g",
	'ĥ': "h", 'ħ': "h",
	'ĩ': "i", 'ī': "i", 'ĭ': "i", 'į': "i", 'ı': "i", 'ĳ': "ij",
	'ĵ': "j", 'ķ': "k",
	'ĺ': "l", 'ļ': "l", 'ľ': "l", 'ŀ': "l", 'ł': "l",
	'ń': "n", 'ņ': "n", 'ň': "n", 'ŉ': "n", 'ŋ': "n",
	'ō': "o", 'ŏ': "o", 'ő': "o", 'œ': "oe",
	'ŕ': "r", 'ŗ': "r", 'ř': "r",
	'ś': "s", 'ŝ': "s", 'ş': "s", 'š': "s", 'ſ': "s",
	'ţ': "t", 'ť': "t", 'ŧ': "t",
	'ũ': "u", 'ū': "u", 'ŭ': "u", 'ů': "u", 'ű': "u", 'ų': "u",
	'ŵ': "w", 'ŷ': "y",
	'ź': "z", 'ż': "z", 'ž': "z",
}

// collationKey returns the primary sort key of a text: one weight for each
// character, with the class in the high bits. Letter case and the marks of
// Latin letters are not in the key.
func collationKey(text string) []uint32 {
	key := make([]uint32, 0, len(text))
	add := func(class int, r rune) { key = append(key, uint32(class)<<24|uint32(r)) }
	for _, r := range strings.ToLower(text) {
		switch {
		case unicode.Is(unicode.Mn, r):
			// A combining mark ("e" + U+0301): the base letter has its
			// own weight already.
		case r >= '0' && r <= '9':
			add(collateDigit, r)
		case unicode.IsLetter(r):
			if base, ok := latinBase[r]; ok {
				for _, b := range base {
					add(collateLetter, b)
				}
			} else {
				add(collateLetter, r)
			}
		case unicode.IsDigit(r):
			add(collateDigit, r)
		default:
			add(collateSymbol, r)
		}
	}
	return key
}

// compareText compares two texts in sort order. It returns 0 only for texts
// that differ at most in letter case, so that a stable sort keeps "Apple"
// and "apple" in their sequence. Texts that differ only in marks ("resume",
// "résumé") are ordered by their lower-case form, so that the result does
// not depend on the sequence of the input.
func compareText(a, b string) int {
	ka, kb := collationKey(a), collationKey(b)
	for i := 0; i < len(ka) && i < len(kb); i++ {
		if ka[i] != kb[i] {
			if ka[i] < kb[i] {
				return -1
			}
			return 1
		}
	}
	if len(ka) != len(kb) {
		if len(ka) < len(kb) {
			return -1
		}
		return 1
	}
	return strings.Compare(strings.ToLower(a), strings.ToLower(b))
}
