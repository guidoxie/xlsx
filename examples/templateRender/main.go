package main

import (
	"github.com/guidoxie/xlsx"
)

type Student struct {
	Name   string
	Course string
	Score  float64
}

type School struct {
	List []Student
}

func main() {
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		panic(err)
	}

	// 单结构体渲染
	if err := f.TemplateRender("Sheet1", &Student{
		Name:   "小明",
		Course: "数学",
		Score:  95.5,
	}); err != nil {
		panic(err)
	}

	// 数组渲染方式一
	school := School{List: []Student{
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
	}}
	if err := f.TemplateRender("Sheet2", school); err != nil {
		panic(err)
	}

	// 数组渲染方式二
	if err := f.TemplateRender("Sheet3", school); err != nil {
		panic(err)
	}
	
	if err := f.SaveAs("template_render.xlsx"); err != nil {
		panic(err)
	}
}
