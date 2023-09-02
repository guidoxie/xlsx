package main

import "github.com/guidoxie/xlsx"

// Student 列宽20，单元格居中
type Student struct {
	Name   string  `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	Course string  `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	Score  float64 `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
}

func main() {
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		panic(err)
	}
	slice := []Student{
		{
			Name:   "小明",
			Course: "数学",
			Score:  95.5,
		},
		{
			Name:   "小红",
			Course: "数学",
			Score:  90.5,
		},
	}
	// SetCursor 设置游标，从第二行开始写
	if err := f.SetCursor("Sheet1", 2).SetRowsValue("Sheet1", slice); err != nil {
		panic(err)
	}
	if err := f.SaveAs("set_rows_value.xlsx"); err != nil {
		panic(err)
	}
}
