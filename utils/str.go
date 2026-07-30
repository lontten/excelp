package utils

import "strings"

// isInvisibleRune 判断是否为应删除的零宽/不可见字符。
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

// mapCleanRune 将单元格中的脏字符映射为删除或替换结果；无需处理时原样返回。
func mapCleanRune(r rune) rune {
	if isInvisibleRune(r) {
		return -1
	}
	switch r {
	case '\u00A0', // NBSP
		'\u202F', // NARROW NO-BREAK SPACE
		'\u205F', // MEDIUM MATHEMATICAL SPACE
		'\u3000': // IDEOGRAPHIC SPACE
		return ' '
	case '\u2028', // LINE SEPARATOR
		'\u2029': // PARAGRAPH SEPARATOR
		return '\n'
	}
	return r
}

func needsClean(s string) bool {
	for _, r := range s {
		if mapCleanRune(r) != r {
			return true
		}
	}
	return false
}

// CleanCell 清洗单元格：删除零宽/不可见字符，将假空格换为普通空格，
// 将行/段分隔符换为换行，再 TrimSpace。正常 \n/\t/\r 保留。
func CleanCell(s string) string {
	if !needsClean(s) {
		return strings.TrimSpace(s)
	}
	return strings.TrimSpace(strings.Map(mapCleanRune, s))
}
