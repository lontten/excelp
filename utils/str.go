package utils

import "strings"

// isInvisibleRune 判断是否为 Excel 常见的零宽/不可见字符。
func isInvisibleRune(r rune) bool {
	switch r {
	case '\u200B', // ZERO WIDTH SPACE
		'\u200C', // ZERO WIDTH NON-JOINER
		'\u200D', // ZERO WIDTH JOINER
		'\uFEFF', // BOM / ZERO WIDTH NO-BREAK SPACE
		'\u2060', // WORD JOINER
		'\u00AD', // SOFT HYPHEN
		'\u200E', // LEFT-TO-RIGHT MARK
		'\u200F': // RIGHT-TO-LEFT MARK
		return true
	}
	return false
}

// CleanCell 去掉单元格中的零宽/不可见字符，再 TrimSpace。
func CleanCell(s string) string {
	needsFilter := false
	for _, r := range s {
		if isInvisibleRune(r) {
			needsFilter = true
			break
		}
	}
	if !needsFilter {
		return strings.TrimSpace(s)
	}
	return strings.TrimSpace(strings.Map(func(r rune) rune {
		if isInvisibleRune(r) {
			return -1
		}
		return r
	}, s))
}
