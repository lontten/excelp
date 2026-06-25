package excelp

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestWrite_SheetIndex(t *testing.T) {
	path := writeTestXLSXMultiSheet(t, []testSheet{
		{name: "Sheet1", rows: [][]string{{"ignored"}}},
		{name: "Sheet2", rows: [][]string{{""}}},
	})

	ctx := ExcelWrite().Template(path).SheetIndex(2)
	if err := Write(ctx, []string{"hello"}); err != nil {
		t.Fatal(err)
	}
	saved, err := ctx.Save()
	if err != nil {
		t.Fatal(err)
	}
	defer os.Remove(saved)

	f, err := excelize.OpenFile(saved)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	val, err := f.GetCellValue("Sheet2", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if val != "hello" {
		t.Errorf("Sheet2 A1 = %q, want %q", val, "hello")
	}
}

func Test_resolveSheetName(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	if _, err := f.NewSheet("Sheet2"); err != nil {
		t.Fatal(err)
	}

	name := "Sheet1"
	idx := 2
	tests := []struct {
		name       string
		sheet      *string
		sheetIndex *int
		want       string
		wantErr    bool
	}{
		{
			name:  "by name",
			sheet: &name,
			want:  "Sheet1",
		},
		{
			name:       "by index",
			sheetIndex: &idx,
			want:       "Sheet2",
		},
		{
			name:    "no sheet",
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := resolveSheetName(f, tt.sheet, tt.sheetIndex)
			if (err != nil) != tt.wantErr {
				t.Fatalf("resolveSheetName() error = %v, wantErr %v", err, tt.wantErr)
			}
			if got != tt.want {
				t.Errorf("resolveSheetName() = %q, want %q", got, tt.want)
			}
		})
	}
}
