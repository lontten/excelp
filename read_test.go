package excelp

import (
	"reflect"
	"testing"
)

func Test_normalizeCol(t *testing.T) {
	tests := []struct {
		name   string
		list   []string
		colNum int
		want   []string
	}{
		{
			name:   "pad when fewer",
			list:   []string{"a", "b"},
			colNum: 3,
			want:   []string{"a", "b", ""},
		},
		{
			name:   "unchanged when equal",
			list:   []string{"a", "b", "c"},
			colNum: 3,
			want:   []string{"a", "b", "c"},
		},
		{
			name:   "truncate when more",
			list:   []string{"a", "b", "c", "d", "e"},
			colNum: 3,
			want:   []string{"a", "b", "c"},
		},
		{
			name:   "no-op when colNum is zero",
			list:   []string{"a", "b", "c"},
			colNum: 0,
			want:   []string{"a", "b", "c"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := normalizeCol(tt.list, tt.colNum)
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("normalizeCol() = %v, want %v", got, tt.want)
			}
		})
	}
}
