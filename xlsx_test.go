package xlsx

import (
	"fmt"
	"testing"
)

func TestFile_SetCellByStruct(t *testing.T) {
	type Student struct {
		Name   string  `xlsx:"axis:A2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Course string  `xlsx:"axis:B2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Score  float64 `xlsx:"axis:C2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	}
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellByStruct("Sheet1", Student{
		Name:   "小明",
		Course: "数学",
		Score:  95.5,
	}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/set_cell_by_struct.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_SetRowValue(t *testing.T) {
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCursor("Sheet1", 2).SetRowValue("Sheet1", []interface{}{"小明", "数学", 95.5}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/set_row_value.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_SetRowsValue(t *testing.T) {
	type Student struct {
		Name   string  `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Course string  `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Score  float64 `xlsx:"colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	}
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
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
	if err := f.SetCursor("Sheet1", 2).SetRowsValue("Sheet1", slice); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/set_rows_value.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_SetRowsValueByTableHeader(t *testing.T) {
	type Student struct {
		Name   string  `xlsx:"column:姓名;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Course string  `xlsx:"column:课程;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Score  float64 `xlsx:"column:成绩;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	}
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
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
	if err := f.SetCursor("Sheet1", 2).SetRowsValueByTableHeader("Sheet1", 1, slice); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/set_rows_value_by_table_header.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_TemplateRender1(t *testing.T) {
	type Student struct {
		Name   string
		Course string
		Score  float64
	}
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.TemplateRender("Sheet2", &Student{
		Name:   "小明",
		Course: "数学",
		Score:  95.5,
	}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/template_render_1.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_TemplateRender2(t *testing.T) {
	type Student struct {
		Name   string
		Course string
		Score  float64
	}
	type School struct {
		List []Student
	}
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
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.TemplateRender("Sheet3", school); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/template_render_2.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_TemplateRender3(t *testing.T) {
	type Student struct {
		Name   string
		Course string
		Score  float64
	}
	type School struct {
		List []Student
	}
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
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.TemplateRender("Sheet4", school); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs("test/template_render_3.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_ReadToSlice(t *testing.T) {
	type Student struct {
		Name   string
		Course string
		Score  float64
	}
	slice := make([]Student, 0)
	f, err := OpenFile("test/template.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	if err := f.ReadToSlice("Sheet5", 2, &slice); err != nil {
		t.Fatal(err)
	}
	fmt.Println(slice[0], slice[1])
}

func TestFile_OpenStreamWriter(t *testing.T) {

}
