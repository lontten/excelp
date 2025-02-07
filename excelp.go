package excelp

import (
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
	"strings"
)

// var ExcelReadContext = ExcelP.Read(url)
type ExcelReadContext struct {
	currentIndex int // excel 的 行标
	url          *string
	sheet        *string
	sheetIndex   *int
	file         *os.File
	excelFile    *excelize.File
	rows         *excelize.Rows
	skip         int
	skipEmptyRow bool // 默认跳过空行
	minCol       int
	err          error

	// ------ 自定义 -----
	convertFunc func(index int, row []string) ([]string, error)

	// ------ T -----
	scanDest any
	scanV    reflect.Value
	scanT    reflect.Type
	m        map[int]Field
}
type Field struct {
	name       string
	timeFormat string
}

func ExcelRead() *ExcelReadContext {
	return &ExcelReadContext{
		skipEmptyRow: true,
	}
}

// initModel 初始化模型
func (c *ExcelReadContext) initModel(dest any) *ExcelReadContext {
	if c.err != nil {
		return c
	}
	v := reflect.ValueOf(dest).Elem()
	t := v.Type()

	c.scanDest = dest
	c.scanV = v
	c.scanT = t
	c.m, c.err = _getStructC(t)
	return c
}

func (c *ExcelReadContext) Url(url string) *ExcelReadContext {
	c.url = &url
	f, err := excelize.OpenFile(url)
	if err != nil {
		c.err = err
		return c
	}
	c.excelFile = f
	return c
}

func (c *ExcelReadContext) Sheet(sheet string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("无 excel file")
		return c
	}
	c.sheet = &sheet
	rows, err := c.excelFile.Rows(sheet)
	if err != nil {
		c.err = err
		return c
	}
	c.rows = rows
	return c
}

// Skip 跳过几行
func (c *ExcelReadContext) Skip(num int) *ExcelReadContext {
	c.skip = num
	return c
}

// SkipEmpty 跳过空行
func (c *ExcelReadContext) SkipEmpty() *ExcelReadContext {
	c.skipEmptyRow = true
	return c
}

// MinCol 设置 最小列数，当列数不足，会填充空字符串
func (c *ExcelReadContext) MinCol(num int) *ExcelReadContext {
	c.minCol = num
	return c
}

// Convert 配置数据转换函数
func (c *ExcelReadContext) Convert(fun func(index int, row []string) ([]string, error)) *ExcelReadContext {
	c.convertFunc = fun
	return c
}

// Close 关闭文件
func (c *ExcelReadContext) Close() error {
	if c.excelFile != nil {
		err := c.excelFile.Close()
		if err != nil {
			return err
		}
	}
	if c.file != nil {
		err := c.file.Close()
		if err != nil {
			return err
		}
	}
	return nil
}

func Read(c *ExcelReadContext, fun func(index int, row []string) error) error {
	return read[int](c, fun, nil)
}

func ReadModel[T any](c *ExcelReadContext, fun func(index int, t T, err *CellErr) error) error {
	return read[T](c, nil, fun)
}

func read[T any](
	c *ExcelReadContext,
	fun1 func(index int, row []string) error,
	fun2 func(index int, t T, err *CellErr) error,
) error {
	if c == nil {
		return errors.New("ExcelReadContext is nil")
	}
	if c.err != nil {
		return c.err
	}

	dest := new(T)
	c.initModel(dest)

	rows := c.rows
	for rows.Next() {
		c.currentIndex++
		if c.currentIndex < c.skip+1 {
			continue
		}
		col, err := rows.Columns()
		if err != nil {
			return err
		}
		var list = make([]string, 0)
		for _, v := range col {
			list = append(list, strings.TrimSpace(v))
		}

		if c.skipEmptyRow {
			if strings.Join(list, "") == "" {
				continue
			}
		}
		if c.minCol > 0 {
			for range c.minCol - len(list) {
				list = append(list, "")
			}
		}

		if c.convertFunc != nil {
			list, err = c.convertFunc(c.currentIndex, list)
			if err != nil {
				return err
			}
		}

		var e error = nil
		if fun1 != nil {
			e = fun1(c.currentIndex, list)
		} else if fun2 != nil {
			t, err := parse[T](c, list)
			e = fun2(c.currentIndex, t, err)
		} else {
			return errors.New("fun1 or fun2 is nil")
		}

		if e != nil {
			if errors.Is(e, ErrExcelPStop) {
				return nil
			}
			return e
		}
	}

	return nil
}
