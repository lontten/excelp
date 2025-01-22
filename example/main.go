package main

import (
	"encoding/json"
	"fmt"
	"github.com/lontten/excelp"
)

func main() {

	readContext := excelp.ExcelRead().
		Url("./excelp_demo.xlsx").
		Sheet("Sheet1").
		MinCol(3).
		Skip(1) //跳过第一行

	defer readContext.Close()

	//excelp.Read(readContext, func(index int, row []string) error {
	//	fmt.Println(index, row)
	//	return nil
	//})

	err := excelp.ReadModel[User](readContext, func(index int, user User, err error) error {
		fmt.Println(index, err)
		bytes, _ := json.Marshal(user)
		fmt.Println(index, string(bytes))
		return nil
	})
	if err != nil {
		fmt.Println(err)
	}
}
