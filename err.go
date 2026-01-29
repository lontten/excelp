package excelp

import (
	"strconv"

	"github.com/pkg/errors"
)

var (
	ErrNil          = errors.New("nil")
	ErrContainEmpty = errors.New("slice empty")

	ErrExcelPStop          = errors.New("excelp stop")            // excelp 停止解析
	ErrExcelPIndexNotFound = errors.New("excelp index not found") // excelp 字段index，找不到
)

type CellErr struct {
	Err   string // 错误信息
	Col   string // 列
	Row   int    // 行
	Value string // excel cell 值
}

func (e CellErr) IsRequiredErr() bool {
	if e.Err == "required" {
		return true
	}
	return false
}

func (e CellErr) ToExcelCellName() string {
	return e.Col + strconv.Itoa(e.Row)
}
