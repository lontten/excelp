package excelp

import (
	"fmt"
	"strings"

	"github.com/lontten/excelp/utils"
	"github.com/lontten/lutil"
	"github.com/pkg/errors"
)

// Read 逐行读取 Excel，对每一行调用 fun。
//
// index 为 Excel 行号，从 1 开始；row 为当前行各列字符串，已去除首尾空格，
// 且若配置了 ColNum 则已做列数归一化（不足补空、超出截断）。
// err 为当前行的单元格级错误列表，可能为空。
//
// 回调返回普通 error 时停止读取并将该错误返回给调用方；
// 回调返回 ErrExcelPStop 时停止读取并返回 nil（正常提前结束）。
func Read(
	c *ExcelReadContext,
	fun func(index int, row []string, err []CellErr) error,
) error {
	return read[int](c, fun, nil)
}

// ReadModel 逐行读取 Excel 并映射到泛型结构体 T，对每一行调用 fun。
//
// 参数语义与 Read 相同；t 为解析后的结构体值。
// 当 err 非空时跳过 parse，t 为零值。
func ReadModel[T any](
	c *ExcelReadContext,
	fun func(index int, row []string, t T, err []CellErr) error,
) error {
	return read[T](c, nil, fun)
}

func panicToError(r any) error {
	switch v := r.(type) {
	case error:
		return v
	case string:
		return errors.New(v)
	default:
		return fmt.Errorf("panic: %v", v)
	}
}

func read[T any](
	c *ExcelReadContext,
	fun1 func(index int, row []string, err []CellErr) error,
	fun2 func(index int, row []string, t T, err []CellErr) error,
) error {
	if c == nil {
		return errors.New("ExcelReadContext is nil")
	}
	c.initRows()
	if err := c.getErr(); err != nil {
		return err
	}
	rows := c.rows
	defer func() {
		if rows != nil {
			_ = rows.Close()
		}
	}()

	var pool *lutil.Pool
	if c.maxLine > 0 {
		pool = lutil.NewPool(c.maxLine, 1, nil)
	}

	if fun2 != nil {
		dest := new(T)
		c.initModel(dest)
		if err := c.getErr(); err != nil {
			return err
		}
	}

	for rows.Next() {
		if c.shouldStop() {
			break
		}
		c.currentIndex++
		index := c.currentIndex
		if index < c.skip+1 {
			continue
		}
		col, err := rows.Columns()
		if err != nil {
			c.setErr(err)
			break
		}

		if pool != nil {
			idx, cols := index, col
			pool.Submit(func() {
				if !c.panic {
					defer func() {
						if r := recover(); r != nil {
							c.setErr(panicToError(r))
						}
					}()
				}

				execErr := doExec(c, idx, fun1, fun2, cols)
				if execErr == nil {
					return
				}
				if errors.Is(execErr, ErrExcelPStop) {
					c.setStop()
					return
				}
				c.setErr(execErr)
			})
		} else {
			execErr := doExec(c, index, fun1, fun2, col)
			if execErr == nil {
				continue
			}
			if errors.Is(execErr, ErrExcelPStop) {
				break
			}
			c.setErr(execErr)
			break
		}
	}

	if pool != nil {
		pool.Shutdown()
	}

	if err := rows.Error(); err != nil {
		c.setErr(err)
	}

	return c.getErr()
}

// normalizeCol 将 list 归一化为固定 colNum 列：不足时末尾补空字符串，超出时截断。
// colNum <= 0 时不做处理，原样返回。
func normalizeCol(list []string, colNum int) []string {
	if colNum <= 0 {
		return list
	}
	if len(list) < colNum {
		for range colNum - len(list) {
			list = append(list, "")
		}
	} else if len(list) > colNum {
		list = list[:colNum]
	}
	return list
}

// doExec 处理单行数据：TrimSpace → SkipEmpty → ColNum → Convert/ConvertCell → parse → 回调。
func doExec[T any](
	c *ExcelReadContext,
	index int,
	fun1 func(index int, row []string, err []CellErr) error,
	fun2 func(index int, row []string, t T, err []CellErr) error,

	col []string) error {
	var err error
	var list = make([]string, 0)
	for _, v := range col {
		list = append(list, strings.TrimSpace(v))
	}

	if c.skipEmptyRow {
		if strings.Join(list, "") == "" {
			return nil
		}
	}
	list = normalizeCol(list, c.colNum)
	var cellErrList = make([]CellErr, 0)

	if c.convertFunc != nil {
		list, err = c.convertFunc(index, list)
		if err != nil {
			cellErrList = append(cellErrList, CellErr{
				Col:   "",
				Err:   err.Error(),
				Row:   index,
				Value: "",
			})
		}
	}

	if len(cellErrList) == 0 {
		for i, f := range c.cellConvertFuncMap {
			if i < 0 || i >= len(list) {
				numberToName, _ := utils.ColumnNumberToName(i)
				cellErrList = append(cellErrList, CellErr{
					Col:   numberToName,
					Err:   "column index out of range",
					Row:   index,
					Value: "",
				})
				continue
			}
			source := list[i]
			target, convErr := f(source)
			if convErr != nil {
				numberToName, _ := utils.ColumnNumberToName(i)
				cellErrList = append(cellErrList, CellErr{
					Col:   numberToName,
					Err:   convErr.Error(),
					Row:   index,
					Value: source,
				})
			}
			list[i] = target
		}
	}

	var e error = nil
	if fun1 != nil {
		e = fun1(index, list, cellErrList)
	} else if fun2 != nil {
		var t T
		if len(cellErrList) == 0 {
			t, cellErrList = parse[T](c, index, list)
		}
		e = fun2(index, list, t, cellErrList)
	} else {
		return errors.New("fun1 or fun2 is nil")
	}

	return e
}
