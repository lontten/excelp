package excelp

import (
	"strings"

	"github.com/lontten/excelp/utils"
	"github.com/lontten/lutil"
	"github.com/pkg/errors"
)

func Read(
	c *ExcelReadContext,
	fun func(index int, row []string, err []CellErr) error,
) error {
	return read[int](c, fun, nil)
}

func ReadModel[T any](
	c *ExcelReadContext,
	fun func(index int, row []string, t T, err []CellErr) error,
) error {
	return read[T](c, nil, fun)
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
	if c.err != nil {
		return c.err
	}
	var pool *lutil.Pool
	if c.maxLine > 0 {
		pool = lutil.NewPool(c.maxLine, 1, nil)
		defer pool.Shutdown()
	}

	if fun2 != nil {
		dest := new(T)
		c.initModel(dest)
	}

	rows := c.rows
	for rows.Next() {
		if c.err != nil {
			return c.err
		}
		c.currentIndex++
		index := c.currentIndex
		if index < c.skip+1 {
			continue
		}
		col, err := rows.Columns()
		if err != nil {
			return err
		}

		if pool != nil {
			pool.Submit(func() {
				if !c.panic {
					defer func() {
						r := recover()
						if r != nil {
							c.err = r.(error)
							return
						}
					}()
				}

				err = doExec(c, index, fun1, fun2, col)
				if err != nil {
					c.err = err
				}
			})
		} else {
			err = doExec(c, index, fun1, fun2, col)
			if err != nil {
				c.err = err
			}
		}
	}
	return nil
}
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
	if c.minCol > 0 {
		for range c.minCol - len(list) {
			list = append(list, "")
		}
	}
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
			source := list[i]
			target, err := f(source)
			if err != nil {
				numberToName, _ := utils.ColumnNumberToName(i)
				cellErrList = append(cellErrList, CellErr{
					Col:   numberToName,
					Err:   err.Error(),
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

	if e != nil {
		if errors.Is(e, ErrExcelPStop) {
			return nil
		}
		return e
	}
	return nil
}
