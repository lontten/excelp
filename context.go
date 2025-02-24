package excelp

import (
	"errors"
	"github.com/lontten/excelp/utils"
	"github.com/lontten/lcore"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
)

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
	convertFunc        func(index int, col []string) ([]string, error)
	cellConvertFuncMap map[int]func(col string) (string, error)

	// ------ T -----
	scanDest any
	scanV    reflect.Value
	scanT    reflect.Type
	m        map[int]Field

	// ------ line -----
	enableAsync bool
	// 启用多线程
	maxLine      int
	waitLine     int
	rejectPolicy lcore.RejectPolicy // 拒绝策略
}
type Field struct {
	name     string
	required bool // 是否必填
}

func ExcelRead() *ExcelReadContext {
	return &ExcelReadContext{
		skipEmptyRow:       true,
		cellConvertFuncMap: make(map[int]func(col string) (string, error)),
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
func (c *ExcelReadContext) initRows() *ExcelReadContext {
	if c.err != nil {
		return c
	}
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	if c.sheet == nil {
		c.err = errors.New("no set sheet")
		return c
	}
	rows, err := c.excelFile.Rows(*c.sheet)
	if err != nil {
		c.err = err
		return c
	}
	c.rows = rows
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
	c.sheet = &sheet
	return c
}

// TimeRow 设置为时间格式
func (c *ExcelReadContext) TimeRow(col ...string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	if c.sheet == nil {
		c.err = errors.New("no set sheet")
		return c
	}
	styleID, _ := c.excelFile.NewStyle(&excelize.Style{NumFmt: 0}) // 常规格式
	for _, s := range col {
		err := c.excelFile.SetColStyle(*c.sheet, s, styleID) // 设置列为常规格式
		if err != nil {
			c.err = err
			return c
		}
	}
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

// EnableAsync 启用异步
func (c *ExcelReadContext) EnableAsync(maxWorkers int, queueSize int, rejectPolicy lcore.RejectPolicy) *ExcelReadContext {
	c.enableAsync = true
	c.maxLine = maxWorkers
	c.waitLine = queueSize
	c.rejectPolicy = rejectPolicy
	return c
}

// ColNum 设置 列数，当列数不足，会填充空字符串
func (c *ExcelReadContext) ColNum(num int) *ExcelReadContext {
	c.minCol = num
	return c
}

// Convert 行转换函数
func (c *ExcelReadContext) Convert(fun func(index int, col []string) ([]string, error)) *ExcelReadContext {
	c.convertFunc = fun
	return c
}

// col 列名，列转化函数
func (c *ExcelReadContext) ConvertCell(col string, fun func(col string) (string, error)) *ExcelReadContext {
	number, err := utils.ColumnNameToNumber(col)
	if err != nil {
		c.err = err
		return c
	}
	c.cellConvertFuncMap[number] = fun
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
