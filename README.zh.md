# excelp


### string
```go

    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        // SheetName("数据").  // 可选，默认第一个工作表
        // SheetIndex(2).     // 可选，按下标指定（从 1 开始）
        ColNum(3).// 固定列数：不足补空、超出截断
        Skip(1) //跳过第一行
    
    defer readContext.Close()
    
    excelp.Read(readContext, func(index int, row []string) error {
        fmt.Println(index, row)
        return nil
    })


```

### struct
```go
    type User struct {
        ID   types.UUID `json:"id"`
        Name string     `json:"info"`
        Age  int        `json:"age"`
    }

    
    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        // SheetName("数据").  // 可选，默认第一个工作表
        ColNum(3).
        Skip(1) //跳过第一行
    
    defer readContext.Close()
    
    err := excelp.ReadModel[User](readContext, func(index int, user *User, err error) error {
		if err != nil {
			return errors.New("read model error")
        }
		bytes, _ := json.Marshal(user)
		fmt.Println(string(bytes))
        return nil
	})
	if err != nil {
        fmt.Println(err)
    }
```
