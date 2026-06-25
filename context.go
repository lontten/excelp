package excelp

import (
	"errors"
	"os"
	"reflect"

	"github.com/lontten/excelp/utils"
	"github.com/lontten/lcore/v2/types"
	"github.com/xuri/excelize/v2"
)

// ExcelReadContext 是 Excel 读取的链式配置上下文，配合 Read 或 ReadModel 使用。
//
// 通过 Url、Sheet、Skip、ColNum 等方法配置后，调用 Read/ReadModel 开始逐行读取。
type ExcelReadContext struct {
	currentIndex int // Excel 行号，从 1 开始递增
	url          *string
	sheet        *string
	sheetIndex   *int
	file         *os.File
	excelFile    *excelize.File
	rows         *excelize.Rows
	skip         int  // 跳过的行数（从第 1 行起计）
	skipEmptyRow bool // 是否跳过全空行，默认 true
	colNum       int  // 固定列数，0 表示不处理
	err          error
	panic        bool // 异步模式下是否允许 panic 向上抛出

	// ------ 自定义 -----
	convertFunc        func(index int, col []string) ([]string, error)
	cellConvertFuncMap map[int]func(col string) (string, error)

	// ------ T -----
	scanDest any
	scanV    reflect.Value
	scanT    reflect.Type
	m        map[int]Field

	// ------ line -----
	maxLine int // 异步 worker 数量，0 表示同步读取
}
type Field struct {
	name     string
	required bool // 是否必填
}

// ExcelRead 创建读取上下文，默认 Sheet 为 Sheet1，默认跳过空行。
func ExcelRead() *ExcelReadContext {
	return &ExcelReadContext{
		sheet:              types.NewString("Sheet1"),
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

// Url 设置 Excel 文件路径并打开文件，支持链式调用。
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

// Sheet 设置要读取的工作表名称，支持链式调用。
func (c *ExcelReadContext) Sheet(sheet string) *ExcelReadContext {
	c.sheet = &sheet
	return c
}

// DateCol 设置为时间格式
// yyyy-MM-dd
func (c *ExcelReadContext) DateCol(col ...string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	if c.sheet == nil {
		c.err = errors.New("no set sheet")
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[1])
	for _, s := range col {
		err := c.excelFile.SetColStyle(*c.sheet, s, styleID)
		if err != nil {
			c.err = err
			return c
		}
	}
	return c
}

// TimeCol 设置为时间格式
// yyyy-MM-dd HH:mm:ss
func (c *ExcelReadContext) TimeCol(col ...string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	if c.sheet == nil {
		c.err = errors.New("no set sheet")
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[2])
	for _, s := range col {
		err := c.excelFile.SetColStyle(*c.sheet, s, styleID)
		if err != nil {
			c.err = err
			return c
		}
	}
	return c
}

// DateTimeCol 设置为时间格式
// yyyy-MM-dd HH:mm:ss
func (c *ExcelReadContext) DateTimeCol(col ...string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	if c.sheet == nil {
		c.err = errors.New("no set sheet")
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[3])
	for _, s := range col {
		err := c.excelFile.SetColStyle(*c.sheet, s, styleID)
		if err != nil {
			c.err = err
			return c
		}
	}
	return c
}

// Skip 设置跳过的行数，从第 1 行起计（常用于跳过表头），支持链式调用。
func (c *ExcelReadContext) Skip(num int) *ExcelReadContext {
	c.skip = num
	return c
}

// Panic 允许异步读取时 panic 向上抛出（默认会 recover），支持链式调用。
func (c *ExcelReadContext) Panic() *ExcelReadContext {
	c.panic = true
	return c
}

// SkipEmpty 启用跳过全空行，支持链式调用。
func (c *ExcelReadContext) SkipEmpty() *ExcelReadContext {
	c.skipEmptyRow = true
	return c
}

// EnableAsync 启用异步读取，maxWorkers 为并发 worker 数量，支持链式调用。
func (c *ExcelReadContext) EnableAsync(maxWorkers int) *ExcelReadContext {
	c.maxLine = maxWorkers
	return c
}

// ColNum 设置每行固定列数，支持链式调用。
//
// 当实际列数少于 num 时，在末尾填充空字符串；
// 当实际列数多于 num 时，截断保留前 num 列；
// num 为 0 或未设置时不做处理。
func (c *ExcelReadContext) ColNum(num int) *ExcelReadContext {
	c.colNum = num
	return c
}

// Convert 设置整行转换函数，在列归一化之后、单元格转换之前执行，支持链式调用。
func (c *ExcelReadContext) Convert(fun func(index int, col []string) ([]string, error)) *ExcelReadContext {
	c.convertFunc = fun
	return c
}

// ConvertCell 为指定列（如 "A"、"B"）设置单元格转换函数，支持链式调用。
func (c *ExcelReadContext) ConvertCell(col string, fun func(col string) (string, error)) *ExcelReadContext {
	number, err := utils.ColumnNameToNumber(col)
	if err != nil {
		c.err = err
		return c
	}
	c.cellConvertFuncMap[number] = fun
	return c
}

// Close 关闭已打开的 Excel 文件及关联资源。
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
