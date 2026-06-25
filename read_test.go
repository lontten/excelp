package excelp

import (
	"fmt"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
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
		{
			name:   "pad empty list",
			list:   []string{},
			colNum: 2,
			want:   []string{"", ""},
		},
		{
			name:   "no-op when colNum is negative",
			list:   []string{"a"},
			colNum: -1,
			want:   []string{"a"},
		},
		{
			name:   "truncate result has exact length",
			list:   []string{"a", "b", "c", "d"},
			colNum: 2,
			want:   []string{"a", "b"},
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

func writeTestXLSX(t *testing.T, sheet string, rows [][]string) string {
	t.Helper()

	f := excelize.NewFile()
	for i, row := range rows {
		for j, cell := range row {
			colName, err := excelize.ColumnNumberToName(j + 1)
			if err != nil {
				t.Fatal(err)
			}
			cellRef := fmt.Sprintf("%s%d", colName, i+1)
			if err := f.SetCellValue(sheet, cellRef, cell); err != nil {
				t.Fatal(err)
			}
		}
	}

	path := filepath.Join(t.TempDir(), "test.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRead_ColNum(t *testing.T) {
	path := writeTestXLSX(t, "Sheet1", [][]string{
		{"a", " b"},
		{"a", "b", "c", "d"},
	})

	ctx := ExcelRead().Url(path).Sheet("Sheet1").ColNum(3)
	defer ctx.Close()

	var got [][]string
	err := Read(ctx, func(index int, row []string, errs []CellErr) error {
		if len(errs) > 0 {
			t.Errorf("row %d: unexpected cell errors: %v", index, errs)
		}
		got = append(got, append([]string(nil), row...))
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}

	want := [][]string{
		{"a", "b", ""},
		{"a", "b", "c"},
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("Read() rows = %v, want %v", got, want)
	}
	for i, row := range got {
		if len(row) != 3 {
			t.Errorf("row %d: len = %d, want 3", i, len(row))
		}
	}
}

type colNumRow struct {
	A string `excelp:"index:0"`
	B string `excelp:"index:1"`
	C string `excelp:"index:2"`
}

func TestReadModel_ColNum(t *testing.T) {
	path := writeTestXLSX(t, "Sheet1", [][]string{
		{"a", "b"},
		{"x", "y", "z", "extra"},
	})

	ctx := ExcelRead().Url(path).Sheet("Sheet1").ColNum(3)
	defer ctx.Close()

	var got []colNumRow
	err := ReadModel(ctx, func(index int, row []string, model colNumRow, errs []CellErr) error {
		if len(errs) > 0 {
			t.Errorf("row %d: unexpected cell errors: %v", index, errs)
		}
		got = append(got, model)
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}

	want := []colNumRow{
		{A: "a", B: "b", C: ""},
		{A: "x", B: "y", C: "z"},
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("ReadModel() models = %v, want %v", got, want)
	}
}
