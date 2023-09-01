package xlsx

import "testing"

func TestFile_SetCellByStruct(t *testing.T) {
	type Student struct {
		Name   string  `xlsx:"axis:A2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Course string  `xlsx:"axis:B2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
		Score  float64 `xlsx:"axis:C2;colWidth:20;style:{\"alignment\":{\"horizontal\":\"center\"}}"`
	}
	f, err := OpenFile("test/set_cell_by_struct.xlsx")
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
	if err := f.SaveAs("test/set_cell_by_struct_result.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func TestFile_SetRowValue(t *testing.T) {

}

func TestFile_SetRowsValue(t *testing.T) {

}

func TestFile_SetRowsValueByTableHeader(t *testing.T) {

}

func TestFile_TemplateRender(t *testing.T) {

}

func TestFile_TemplateRenderAllSheet(t *testing.T) {

}

func TestFile_ReadToSlice(t *testing.T) {

}
