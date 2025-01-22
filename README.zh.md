# excelp


### string
```go

    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        Sheet("Sheet1").
        MinCol(3).// 给row 设置最小列数，填充默认空字符
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
        ID   types.UUID `json:"id"  tableName:"public.t_user"`
        Name string     `json:"info"`
        Age  int        `json:"age"`
    }

    
    readContext := excelp.ExcelRead().
        Url("./excelp_demo.xlsx").
        Sheet("Sheet1").
        MinCol(3).
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
