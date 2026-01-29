package utils

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestColumnNameToNumber(t *testing.T) {

	got, err := excelize.ColumnNameToNumber("a")
	fmt.Println(got, err)

}
