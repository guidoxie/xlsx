## 概述
对 [Excelize](https://github.com/qax-os/excelize) 进行了简单的封装，方便的读取Excel文件和Excel文件里面写入数据，大部分情况下，定义好结构体就行， 而不需要在代码里面指定各种坐标，让代码难以维护和阅读
> 行和列从1开始算

## 字段标签
| 标签名      | 说明 |
| ----------- | ----------- |
| axis      |指定单元格坐标，不指定时，默认为字段在结构体的顺序，如为结构体的第一个字段，则列的坐标为A，第一列|
| style   | 指定单元格样式，excelize.Style{}序列化后的json字符串，参考文档：https://xuri.me/excelize/zh-hans/style.html |
| column   | 指定列名，用于根据表头的和列名来匹配对应的列坐标|
| colWidth   | 指定列的宽度 |

## 示例
### 写入一个结构体数据
```go
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
```

### 写一行数据
```go
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
```
### 写入多行数据
```go
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
```
### 根据表头写入多行
```go
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
```
### 模板语法
```go
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
```
### 读取多行数据
```go
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
	t.Log(slice)
}
```

## 参考链接
* Excelize文档：https://xuri.me/excelize/zh-hans/