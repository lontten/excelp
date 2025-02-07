package excelp

import "github.com/pkg/errors"

var (
	ErrNil          = errors.New("nil")
	ErrContainEmpty = errors.New("slice empty")

	ErrExcelPStop          = errors.New("excelp stop")            // excelp 停止解析
	ErrExcelPIndexNotFound = errors.New("excelp index not found") // excelp 字段index，找不到
)

type CellErr struct {
	Err   string
	Col   int
	Value string
}
