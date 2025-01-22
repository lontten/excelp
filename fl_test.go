package excelp

import (
	"fmt"
	"github.com/gofrs/uuid"
	"github.com/shopspring/decimal"
	"github.com/xuri/excelize/v2"
	"reflect"
	"testing"
	"time"
)

type User struct {
	ID       *int64           `excelp:"index:0"`
	Date     *time.Time       `excelp:"index:1;format:2006-01-02;"`
	Time     *time.Time       `excelp:"index:2;format:15:04:05;"`
	DateTime *time.Time       `excelp:"index:3;format:2006-01-02 15:04:05;"`
	Uuid     *uuid.UUID       `excelp:"index:4"`
	Money    *decimal.Decimal `excelp:"index:5"`
	Fl       *float64         `excelp:"index:g"`
	Name     *string          `excelp:"index:h"`

	Date2     *time.Time `excelp:"index:i;format:2006-01-02;"`
	Time2     *time.Time `excelp:"index:j;format:15:04:05;"`
	DateTime2 *time.Time `excelp:"index:k;format:2006-01-02 15:04:05;"`
}

func Test__getStructC(t *testing.T) {
	user := User{}
	of := reflect.TypeOf(user)

	gotM, err := _getStructC(of)
	if err != nil {
		t.Errorf("_getStructC() error = %v", err)
		return
	}
	for i, field := range gotM {
		fmt.Println(i, field)
	}
}

func Test__ColumnNumberToName(t *testing.T) {
	name, err := excelize.ColumnNumberToName(6)
	fmt.Println(name, err)

	number, err := excelize.ColumnNameToNumber("g")
	fmt.Println(number, err)
}
