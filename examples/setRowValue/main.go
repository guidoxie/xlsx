package main

import "github.com/guidoxie/xlsx"

func main() {
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		panic(err)
	}
	// SetCursor 设置游标，从第二行开始写
	if err := f.SetCursor("Sheet1", 2).SetRowValue("Sheet1", []interface{}{"小明", "数学", 95.5}); err != nil {
		panic(err)
	}
	if err := f.SaveAs("set_row_value.xlsx"); err != nil {
		panic(err)
	}
}
