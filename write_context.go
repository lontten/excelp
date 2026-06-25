package excelp

import (
	"errors"
	"fmt"
	"os"
	"reflect"

	"github.com/lontten/lcore/v2/types"
	"github.com/lontten/lutil/fileutil"
	"github.com/xuri/excelize/v2"
)

type ExcelWriteContext struct {
	template *string

	sheet      *string
	sheetIndex *int
	file       *os.File
	excelFile  *excelize.File

	err          error
	currentIndex int // excel 的 行标

	// ------ 自定义 -----
	convertFunc        func(index int, col []string) ([]string, error)
	cellConvertFuncMap map[int]func(col string) (string, error)

	// ------ T -----
	scanDest any
	scanV    reflect.Value
	scanT    reflect.Type
	m        map[int]Field
}

func ExcelWrite() *ExcelWriteContext {
	return &ExcelWriteContext{
		sheet:              types.NewString("Sheet1"),
		cellConvertFuncMap: make(map[int]func(col string) (string, error)),
	}
}

// initModel 初始化模型
func (c *ExcelWriteContext) initModel(dest any) *ExcelWriteContext {
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

func (c *ExcelWriteContext) Template(path string) *ExcelWriteContext {
	c.template = &path
	file, err := fileutil.CopyTemplateToTempFile(path)
	if err != nil {
		c.err = err
		return c
	}
	c.file = file
	f, err := excelize.OpenFile(file.Name())
	if err != nil {
		c.err = err
		return c
	}
	c.excelFile = f
	return c
}

// SheetName 按名称设置要写入的工作表，与 SheetIndex 互斥，后调用者生效。
func (c *ExcelWriteContext) SheetName(name string) *ExcelWriteContext {
	c.sheet = &name
	c.sheetIndex = nil
	return c
}

// SheetIndex 按下标设置要写入的工作表，下标从 1 开始（1 表示第一个工作表），与 SheetName 互斥。
//
// 需在 Template 打开文件后才能解析实际工作表名称。
func (c *ExcelWriteContext) SheetIndex(index int) *ExcelWriteContext {
	c.sheetIndex = &index
	c.sheet = nil
	return c
}

func (c *ExcelWriteContext) sheetName() (string, error) {
	if c.err != nil {
		return "", c.err
	}
	return resolveSheetName(c.excelFile, c.sheet, c.sheetIndex)
}

// Skip 跳过几行
func (c *ExcelWriteContext) Skip(num int) *ExcelWriteContext {
	c.currentIndex = num
	return c
}

// Save 关闭文件,返回文件路径
func (c *ExcelWriteContext) Save() (string, error) {
	if c.excelFile == nil {
		return "", errors.New("template err")
	}
	err := c.excelFile.Save()
	if err != nil {
		return "", fmt.Errorf("保存excelFile失败: %w", err)
	}
	err = c.excelFile.Close()
	if err != nil {
		return "", fmt.Errorf("关闭excelFile失败: %w", err)
	}
	return c.file.Name(), nil
}

// Remove 删除临时文件
func (c *ExcelWriteContext) Remove() error {
	return os.Remove(c.file.Name())
}
