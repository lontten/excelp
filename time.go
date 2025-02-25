package excelp

import (
	"github.com/lontten/lcore/types"
	"github.com/xuri/excelize/v2"
)

var TimeFormat = map[int]*excelize.Style{
	0: {NumFmt: 0},
	1: {CustomNumFmt: types.NewString("yyyy-mm-dd")},
	2: {CustomNumFmt: types.NewString("HH:mm:ss")},
	3: {CustomNumFmt: types.NewString("yyyy-mm-dd HH:mm:ss")},
}
