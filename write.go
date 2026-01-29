package excelp

import (
	"strconv"

	"github.com/lontten/excelp/utils"
	"github.com/pkg/errors"
)

func WriteModel[T any](t T) error {

	return nil
}
func Write(c *ExcelWriteContext, col []string) error {
	if c == nil {
		return errors.New("ExcelWriteContext is nil")
	}
	if c.err != nil {
		return c.err
	}
	if c.excelFile == nil {
		return errors.New("template err")
	}
	c.currentIndex++
	for i, s := range col {
		name, _ := utils.ColumnNumberToName(i)
		err := c.excelFile.SetCellValue(*c.sheet, name+strconv.Itoa(c.currentIndex), s)
		if err != nil {
			return err
		}
	}
	return nil
}
