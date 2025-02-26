package excelp

import (
	"github.com/lontten/excelp/utils"
	"github.com/pkg/errors"
	"strconv"
)

func WriteModel[T any](t T) error {

	return nil
}
func Write(c *ExcelWriteContext, col []string) error {
	if c == nil {
		return errors.New("ExcelWriteContext is nil")
	}

	c.mu.Lock()
	defer c.mu.Unlock()

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
