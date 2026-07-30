package utils

import "testing"

func TestCleanCell(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{
			name: "trim ordinary spaces",
			in:   "  hello  ",
			want: "hello",
		},
		{
			name: "unchanged normal text",
			in:   "业务数据123",
			want: "业务数据123",
		},
		{
			name: "strip leading and trailing zero-width",
			in:   "\u200B\uFEFFhello\u200C\u2060",
			want: "hello",
		},
		{
			name: "strip embedded zero-width",
			in:   "hel\u200Blo\u200Dworld",
			want: "helloworld",
		},
		{
			name: "only zero-width becomes empty",
			in:   "\u200B\u200C\u200D\uFEFF\u2060\u00AD\u200E\u200F",
			want: "",
		},
		{
			name: "soft hyphen and direction marks",
			in:   "\u200Ea\u00ADb\u200F",
			want: "ab",
		},
		{
			name: "trim after removing zero-width",
			in:   " \u200B hi \uFEFF ",
			want: "hi",
		},
		{
			name: "empty stays empty",
			in:   "",
			want: "",
		},
		{
			name: "NBSP middle becomes space",
			in:   "LINE\u00A01-4",
			want: "LINE 1-4",
		},
		{
			name: "NBSP edges trimmed after replace",
			in:   "\u00A0hello\u00A0",
			want: "hello",
		},
		{
			name: "ideographic and narrow spaces",
			in:   "a\u3000b\u202Fc\u205Fd",
			want: "a b c d",
		},
		{
			name: "paragraph separator to newline",
			in:   "LSQWRF100/R2Y\u2029名义制冷量",
			want: "LSQWRF100/R2Y\n名义制冷量",
		},
		{
			name: "line separator to newline",
			in:   "a\u2028b",
			want: "a\nb",
		},
		{
			name: "keep normal newline and tab",
			in:   "a\nb\tc",
			want: "a\nb\tc",
		},
		{
			name: "city with trailing ZWNJ",
			in:   "湖北省，武汉市\u200C",
			want: "湖北省，武汉市",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := CleanCell(tt.in)
			if got != tt.want {
				t.Errorf("CleanCell(%q) = %q, want %q", tt.in, got, tt.want)
			}
		})
	}
}
