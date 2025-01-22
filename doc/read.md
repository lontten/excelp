```go



var url = "xx.xlsx"
var ExcelRead = ExcelP.Read(url)
ExcelRead.skip(2)
ExcelRead.Read(func (index int,row []string) {

})


while ExcelRead.Next() {
    ExcelP.Read(func (index int,row []string) {
            fmt.Println(index,row)
	})
}


while ExcelRead.Next() {
    ExcelP.ReadModel[User](ExcelRead, func (index int, row User) {
    
    })
}


```