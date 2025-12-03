package main

import (
	"encoding/json"
	"fmt"

	"github.com/lontten/excelp"
	"github.com/lontten/lcore"
)

func main() {

	readContext := excelp.ExcelRead().
		Url("./excelp_demo.xlsx").
		Sheet("Sheet1").
		EnableAsync(2, 2, lcore.CallerRunsPolicy).
		ColNum(3).
		Skip(1) //跳过第一行

	defer readContext.Close()

	//excelp.Read(readContext, func(index int, row []string) error {
	//	fmt.Println(index, row)
	//	return nil
	//})

	err := excelp.ReadModel[User](readContext, func(index int, row []string, user User, e []excelp.CellErr) error {
		if len(e) > 0 {
			bytes, _ := json.Marshal(e)
			fmt.Println(index, string(bytes))
		} else {
			bytes, _ := json.Marshal(user)
			fmt.Println(index, string(bytes))
		}
		return nil
	})
	if err != nil {
		fmt.Println(err)
	}
}
