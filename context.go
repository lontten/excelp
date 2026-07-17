package excelp

import (
	"errors"
	"fmt"
	"os"
	"reflect"
	"sync"

	"github.com/lontten/excelp/utils"
	"github.com/xuri/excelize/v2"
)

// ExcelReadContext 是 Excel 读取的链式配置上下文，配合 Read 或 ReadModel 使用。
//
// 通过 Url、SheetName、SheetIndex、Skip、ColNum 等方法配置后，调用 Read/ReadModel 开始逐行读取。
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
	errMu        sync.Mutex
	err          error
	stop         bool // 收到 ErrExcelPStop 后停止继续提交/读取
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

func (c *ExcelReadContext) setErr(err error) {
	if err == nil {
		return
	}
	c.errMu.Lock()
	defer c.errMu.Unlock()
	if c.err == nil {
		c.err = err
	}
}

func (c *ExcelReadContext) getErr() error {
	c.errMu.Lock()
	defer c.errMu.Unlock()
	return c.err
}

func (c *ExcelReadContext) setStop() {
	c.errMu.Lock()
	defer c.errMu.Unlock()
	c.stop = true
}

func (c *ExcelReadContext) shouldStop() bool {
	c.errMu.Lock()
	defer c.errMu.Unlock()
	return c.stop || c.err != nil
}

type Field struct {
	name     string
	required bool // 是否必填
}

// ExcelRead 创建读取上下文，未指定工作表时默认使用第一个，默认跳过空行。
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
	sheetName, err := c.sheetName()
	if err != nil {
		c.err = err
		return c
	}
	rows, err := c.excelFile.Rows(sheetName)
	if err != nil {
		c.err = err
		return c
	}
	c.rows = rows
	return c
}

func resolveSheetName(f *excelize.File, sheet *string, sheetIndex *int) (string, error) {
	if f == nil {
		return "", errors.New("no set excel")
	}
	if sheet != nil {
		return *sheet, nil
	}
	if sheetIndex != nil {
		list := f.GetSheetList()
		idx := *sheetIndex
		if idx < 1 || idx > len(list) {
			return "", fmt.Errorf("sheet index %d out of range, total %d", idx, len(list))
		}
		return list[idx-1], nil
	}
	list := f.GetSheetList()
	if len(list) == 0 {
		return "", errors.New("no sheet in workbook")
	}
	return list[0], nil
}

func (c *ExcelReadContext) sheetName() (string, error) {
	if c.err != nil {
		return "", c.err
	}
	return resolveSheetName(c.excelFile, c.sheet, c.sheetIndex)
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

// SheetName 按名称设置要读取的工作表，与 SheetIndex 互斥，后调用者生效。
// 若均未设置，默认使用第一个工作表。
func (c *ExcelReadContext) SheetName(name string) *ExcelReadContext {
	c.sheet = &name
	c.sheetIndex = nil
	return c
}

// SheetIndex 按下标设置要读取的工作表，下标从 1 开始（1 表示第一个工作表），与 SheetName 互斥。
// 若均未设置，默认使用第一个工作表。
//
// 需在 Url 打开文件后才能解析实际工作表名称。
func (c *ExcelReadContext) SheetIndex(index int) *ExcelReadContext {
	c.sheetIndex = &index
	c.sheet = nil
	return c
}

// DateCol 设置为时间格式
// yyyy-MM-dd
func (c *ExcelReadContext) DateCol(col ...string) *ExcelReadContext {
	if c.excelFile == nil {
		c.err = errors.New("no set excel")
		return c
	}
	sheetName, err := c.sheetName()
	if err != nil {
		c.err = err
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[1])
	for _, s := range col {
		err := c.excelFile.SetColStyle(sheetName, s, styleID)
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
	sheetName, err := c.sheetName()
	if err != nil {
		c.err = err
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[2])
	for _, s := range col {
		err := c.excelFile.SetColStyle(sheetName, s, styleID)
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
	sheetName, err := c.sheetName()
	if err != nil {
		c.err = err
		return c
	}

	styleID, _ := c.excelFile.NewStyle(TimeFormat[3])
	for _, s := range col {
		err := c.excelFile.SetColStyle(sheetName, s, styleID)
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
