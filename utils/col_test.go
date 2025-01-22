package utils

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestColumnNameToNumber(t *testing.T) {

	got, err := excelize.ColumnNameToNumber("a")
	fmt.Println(got, err)

}
