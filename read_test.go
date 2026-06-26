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

func writeRowsToSheet(t *testing.T, f *excelize.File, sheet string, rows [][]string) {
	t.Helper()
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
}

type testSheet struct {
	name string
	rows [][]string
}

func writeTestXLSXMultiSheet(t *testing.T, sheets []testSheet) string {
	t.Helper()
	if len(sheets) == 0 {
		t.Fatal("no sheets")
	}

	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", sheets[0].name); err != nil {
		t.Fatal(err)
	}
	writeRowsToSheet(t, f, sheets[0].name, sheets[0].rows)
	for i := 1; i < len(sheets); i++ {
		if _, err := f.NewSheet(sheets[i].name); err != nil {
			t.Fatal(err)
		}
		writeRowsToSheet(t, f, sheets[i].name, sheets[i].rows)
	}

	path := filepath.Join(t.TempDir(), "multi.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func writeTestXLSX(t *testing.T, sheet string, rows [][]string) string {
	t.Helper()

	f := excelize.NewFile()
	writeRowsToSheet(t, f, sheet, rows)

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

	ctx := ExcelRead().Url(path).SheetName("Sheet1").ColNum(3)
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

	ctx := ExcelRead().Url(path).SheetName("Sheet1").ColNum(3)
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

func readAllRows(t *testing.T, ctx *ExcelReadContext) [][]string {
	t.Helper()
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
	return got
}

func TestRead_SheetByName(t *testing.T) {
	path := writeTestXLSXMultiSheet(t, []testSheet{
		{name: "Sheet1", rows: [][]string{{"only-sheet1"}}},
		{name: "Sheet2", rows: [][]string{{"target-data"}}},
	})

	ctx := ExcelRead().Url(path).SheetName("Sheet2")
	defer ctx.Close()

	got := readAllRows(t, ctx)
	want := [][]string{{"target-data"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("Read() rows = %v, want %v", got, want)
	}
}

func TestRead_SheetByIndex(t *testing.T) {
	path := writeTestXLSXMultiSheet(t, []testSheet{
		{name: "Sheet1", rows: [][]string{{"only-sheet1"}}},
		{name: "Sheet2", rows: [][]string{{"target-data"}}},
	})

	ctx := ExcelRead().Url(path).SheetIndex(2)
	defer ctx.Close()

	got := readAllRows(t, ctx)
	want := [][]string{{"target-data"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("Read() rows = %v, want %v", got, want)
	}
}

func TestRead_SheetIndex_outOfRange(t *testing.T) {
	path := writeTestXLSX(t, "Sheet1", [][]string{{"a"}})

	ctx := ExcelRead().Url(path).SheetIndex(99)
	defer ctx.Close()

	err := Read(ctx, func(index int, row []string, errs []CellErr) error {
		return nil
	})
	if err == nil {
		t.Fatal("expected error for out of range sheet index")
	}
}

func TestRead_SheetIndex_zero(t *testing.T) {
	path := writeTestXLSX(t, "Sheet1", [][]string{{"a"}})

	ctx := ExcelRead().Url(path).SheetIndex(0)
	defer ctx.Close()

	err := Read(ctx, func(index int, row []string, errs []CellErr) error {
		return nil
	})
	if err == nil {
		t.Fatal("expected error for sheet index 0")
	}
}

func TestRead_DefaultFirstSheet(t *testing.T) {
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "CustomFirst"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("CustomFirst", "A1", "first-sheet-data"); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "default.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}

	ctx := ExcelRead().Url(path)
	defer ctx.Close()

	got := readAllRows(t, ctx)
	want := [][]string{{"first-sheet-data"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("Read() rows = %v, want %v", got, want)
	}
}
