# excelp

### string

```

    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        Sheet("Sheet1").
        ColNum(3).// 固定列数：不足补空、超出截断
        Skip(1) //跳过第一行
    
    defer readContext.Close()
    
    excelp.Read(readContext, func(index int, row []string) error {
        fmt.Println(index, row)
        return nil
    })


```

### struct

```
    type User struct {
        ID   types.UUID `json:"id""`
        Name string     `json:"info"`
        Age  int        `json:"age"`
    }

    
    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        Sheet("Sheet1").
        ColNum(3).
        Skip(1) //跳过第一行
    
    defer readContext.Close()
    
	err := excelp.ReadModel[User](readContext, func(index int, row []string, user User, e []excelp.CellErr) error {
		if len(e) > 0 {
			bytes, _ := json.Marshal(e)
			fmt.Println(index, string(bytes))
		} else {
			bytes, _ := json.Marshal(user)
			fmt.Println(index, string(bytes))
		}
		return nil //这里返回err，ReadModel会在读取新一行前，检查err并返回
	})
	if err != nil {
		fmt.Println(err)
	}
```
