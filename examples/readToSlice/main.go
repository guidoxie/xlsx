package main

import (
	"fmt"
	"github.com/guidoxie/xlsx"
)

type Student struct {
	Name   string
	Course string
	Score  float64
}

func main() {
	slice := make([]Student, 0)
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		panic(err)
	}
	// 从第二开始读取
	if err := f.ReadToSlice("Sheet1", 2, &slice); err != nil {
		panic(err)
	}
	fmt.Println(slice)
}
