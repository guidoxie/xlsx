package main

import "github.com/guidoxie/xlsx"

// Student Name写到A1, Course写到B2， Score写到C2列宽20，单元格居中
type Student struct {
	Name   string  `xlsx:"axis:A2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	Course string  `xlsx:"axis:B2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	Score  float64 `xlsx:"axis:C2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
}

func main() {
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		panic(err)
	}
	if err := f.SetCellByStruct("Sheet1", Student{
		Name:   "小明",
		Course: "数学",
		Score:  95.5,
	}); err != nil {
		panic(err)
	}
	if err := f.SaveAs("set_cell_by_struct.xlsx"); err != nil {
		panic(err)
	}
}
