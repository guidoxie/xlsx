package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
	"io"
	"reflect"
	"regexp"
	"strings"
	"sync"
	"text/template"
)

type Picture []byte // 图片

type File struct {
	*excelize.File
	cursor map[string]int           // 游标，写到那一行了
	writer map[string]*StreamWriter // 流式写入器
	mx     sync.Mutex
}

type StreamWriter struct {
	*excelize.StreamWriter // 流式写入器
}

func NewFile() *File {
	return &File{
		File:   excelize.NewFile(),
		cursor: make(map[string]int),
		writer: make(map[string]*StreamWriter),
		mx:     sync.Mutex{},
	}
}

func OpenFile(filename string) (*File, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	return &File{
		File:   f,
		cursor: make(map[string]int),
		writer: make(map[string]*StreamWriter),
		mx:     sync.Mutex{},
	}, nil
}

func OpenReader(r io.Reader) (*File, error) {
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}
	return &File{
		File:   f,
		cursor: make(map[string]int),
		writer: make(map[string]*StreamWriter),
		mx:     sync.Mutex{},
	}, nil
}

func OpenBytes(b []byte) (*File, error) {
	return OpenReader(bytes.NewBuffer(b))
}

func (f *File) CloseStreamWriter(sheet ...string) error {
	if len(sheet) > 0 {
		for _, s := range sheet {
			sw, ok := f.writer[s]
			if !ok {
				continue
			}
			if err := sw.Flush(); err != nil {
				return err
			}
			delete(f.writer, s)
		}
	} else {
		for s, sw := range f.writer {
			if err := sw.Flush(); err != nil {
				return err
			}
			delete(f.writer, s)
		}
	}
	return nil
}

func (f *File) OpenStreamWriter(sheet string) (*StreamWriter, error) {
	f.mx.Lock()
	defer f.mx.Unlock()
	if sw, ok := f.writer[sheet]; ok {
		return sw, nil
	}
	sw, err := f.NewStreamWriter(sheet)
	if err != nil {

		return nil, err
	}
	f.writer[sheet] = &StreamWriter{sw}
	return f.writer[sheet], nil
}

func (f *File) getStreamWriter(sheet string) *StreamWriter {
	return f.writer[sheet]
}

func (f *File) SetCursor(sheet string, cursor int) *File {
	f.mx.Lock()
	defer f.mx.Unlock()
	if cursor > 0 {
		f.cursor[sheet] = cursor
	}
	return f
}

func (f *File) GetCursor(sheet string) int {
	f.mx.Lock()
	defer f.mx.Unlock()
	if _, ok := f.cursor[sheet]; !ok {
		f.cursor[sheet] = 1
	}
	return f.cursor[sheet]
}

// 自增
func (f *File) cursorInc(sheet string) int {
	f.mx.Lock()
	defer f.mx.Unlock()
	if _, ok := f.cursor[sheet]; !ok {
		f.cursor[sheet] = 1
	}
	f.cursor[sheet] = f.cursor[sheet] + 1
	return f.cursor[sheet]
}

// SetCellByStruct 根据结构体设置单元格值
func (f *File) SetCellByStruct(sheet string, dataStruct interface{}) error {
	return f.setCellByStruct(sheet, reflect.ValueOf(dataStruct), nil)
}

func (f *File) setCellByStruct(sheet string, data reflect.Value, tableHeader map[string]string) error {
	var (
		typeOf  = data.Type()
		valueOf = data
	)
	if typeOf.Kind() == reflect.Ptr {
		typeOf = typeOf.Elem()
		valueOf = valueOf.Elem()
	}

	for i := 0; i < typeOf.NumField(); i++ {
		tag, err := ParseTagSetting(typeOf.Field(i), i, f.GetCursor(sheet))
		if err != nil {
			return err
		}
		if tag.Ignore {
			continue
		}
		var axis string
		c, ok := tableHeader[tag.Column]
		if ok {
			axis = fmt.Sprintf("%s%d", c, f.GetCursor(sheet))
		} else if len(tableHeader) == 0 {
			cells := strings.Split(tag.Axis, "-")
			// 合并单元格
			if len(cells) == 2 {
				err = f.MergeCell(sheet, cells[0], cells[1])
			}
			if err != nil {
				return err
			}
			axis = cells[0]
		}
		if len(axis) == 0 {
			continue
		}
		cell := &excelize.Cell{
			Value: valueOf.Field(i).Interface(),
		}
		// 设置单元格样式
		if tag.Style != nil {
			sid, err := f.NewStyle(tag.Style)
			if err != nil {
				return err
			}
			cell.StyleID = sid
		}
		err = f.setCell(sheet, axis, cell)
		if err != nil {
			return err
		}
		// 设置列宽度
		if tag.ColWidth > 0 {
			if err := f.SetColWidth(sheet, tag.Col, tag.Col, tag.ColWidth); err != nil {
				return err
			}
		}
	}
	return nil
}

// SetRowsValue 设置多行的值，支持 1.多维度数组或切片 2.结构体数组或切片
func (f *File) SetRowsValue(sheet string, slice interface{}) error {
	valueOf := reflect.ValueOf(slice)
	for i := 0; i < valueOf.Len(); i++ {
		data := valueOf.Index(i)
		var (
			dataTypeOf  = data.Type()
			dataValueOf = data
		)
		if dataTypeOf.Kind() == reflect.Ptr {
			dataTypeOf = dataTypeOf.Elem()
			dataValueOf = dataValueOf.Elem()
		}
		switch dataTypeOf.Kind() {
		case reflect.Struct:
			if err := f.setCellByStruct(sheet, dataValueOf, nil); err != nil {
				return err
			}
		case reflect.Array, reflect.Slice:
			if err := f.setRowValue(sheet, dataValueOf); err != nil {
				return err
			}
		}
		f.cursorInc(sheet)
	}
	return nil
}

// SetRowValue 设置一行的值，支持数组和切片
func (f *File) SetRowValue(sheet string, slice interface{}) error {
	if err := f.setRowValue(sheet, reflect.ValueOf(slice)); err != nil {
		return err
	}
	f.cursorInc(sheet)
	return nil
}

func (f *File) setRowValue(sheet string, valueOf reflect.Value) error {
	for i := 0; i < valueOf.Len(); i++ {
		data := valueOf.Index(i)
		var (
			dataTypeOf  = data.Type()
			dataValueOf = data
		)
		if dataTypeOf.Kind() == reflect.Ptr {
			dataTypeOf = dataTypeOf.Elem()
			dataValueOf = dataValueOf.Elem()
		}
		err := f.setCell(sheet, GetAxis(i+1, f.GetCursor(sheet)), &excelize.Cell{Value: dataValueOf.Interface()})
		if err != nil {
			return err
		}
	}
	return nil
}

// SetRowsValueByTableHeader 根据表头设置多行数据
func (f *File) SetRowsValueByTableHeader(sheet string, tableHeader map[string]string, slice interface{}) error {
	valueOf := reflect.ValueOf(slice)
	for i := 0; i < valueOf.Len(); i++ {
		if err := f.setCellByStruct(sheet, valueOf.Index(i), tableHeader); err != nil {
			return err
		}
		f.cursorInc(sheet)
	}
	return nil
}

func (f *File) getTableHeader(sheet string, tableHeader int) (map[string]string, error) {
	// 获取列名
	column := make(map[string]string)
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}
	var count = 1
	for rows.Next() {
		if count == tableHeader {
			row, err := rows.Columns()
			if err != nil {
				return nil, err
			}
			for i, colCell := range row {
				column[strings.TrimSpace(colCell)] = GetAxis(i + 1)
			}
			break
		}
		count++
	}
	if err := rows.Close(); err != nil {
		return nil, err
	}
	return column, nil
}

// TemplateRenderAllSheet 模板渲染遍历所有sheet，注意结构体标签会失效
func (f *File) TemplateRenderAllSheet(data interface{}) error {
	for _, sheet := range f.GetSheetList() {
		if err := f.TemplateRender(sheet, data); err != nil {
			return err
		}
	}
	return nil
}

// TemplateRender 模板渲染，注意结构体标签会失效
func (f *File) TemplateRender(sheet string, data interface{}) error {
	if err := f.beforeTemplateRender(sheet, data); err != nil {
		return err
	}

	// 渲染值
	rows, err := f.File.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	rg := regexp.MustCompile("{{.*}}")
	rangeStartRg := regexp.MustCompile(`{{\s?range \.(.*?)\s?}}`)
	rangeEndRg := regexp.MustCompile(`{{\s?end\s?}}`)
	renderData := data
	rangeStart := false
	delRow := make([]int, 0)
	var rangeArray reflect.Value
	var rangeArrayIndex int
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return err
		}
		if rangeStart && rangeArrayIndex < rangeArray.Len() {
			renderData = rangeArray.Index(rangeArrayIndex).Interface()
			rangeArrayIndex++
		}
		for i, colCell := range row {
			if rangeStartRg.MatchString(colCell) {
				rangeStart = true
				delRow = append(delRow, f.GetCursor(sheet))
				sub := rangeStartRg.FindStringSubmatch(colCell)
				v := reflect.ValueOf(data)
				if v.Kind() == reflect.Ptr {
					v = reflect.ValueOf(data).Elem()
				}
				var temp = v
				for _, n := range strings.Split(sub[1], ".") {
					rangeArray = temp.FieldByName(strings.TrimSpace(n))
					temp = rangeArray
				}
				if rangeArray.Len() == 0 {
					delRow = append(delRow, f.GetCursor(sheet)+1)
				}
				break
			}
			if rangeEndRg.MatchString(colCell) {
				rangeStart = false
				delRow = append(delRow, f.GetCursor(sheet))
				renderData = data
				rangeArrayIndex = 0
				break
			}
			// 语法校验
			if !rg.MatchString(colCell) {
				continue
			}
			buffer := bytes.NewBufferString("")
			tmpl, err := template.New(sheet).Parse(colCell)
			if err != nil {
				return err
			}
			if err = tmpl.Execute(buffer, renderData); err != nil {
				return err
			}
			newCell := buffer.String()
			if colCell != newCell {
				// TODO 转换成对应的数据类型
				number, err := cast.ToFloat64E(newCell)
				if err == nil { // 数字类型
					err = f.SetCellValue(sheet, GetAxis(i+1, f.GetCursor(sheet)), number)
				} else {
					err = f.SetCellValue(sheet, GetAxis(i+1, f.GetCursor(sheet)), newCell)
				}
				if err != nil {
					return err
				}
			}
		}
		f.cursorInc(sheet)
	}
	for i := len(delRow) - 1; i >= 0; i-- {
		if err := f.File.RemoveRow(sheet, delRow[i]); err != nil {
			return err
		}
	}
	// 处理公式
	if err := f.File.UpdateLinkedValue(); err != nil {
		return err
	}
	return nil
}

// ReadToSlice 读取表格数据放至slice中
//
// startRow: 开始行
//
// endRow: 可选参数，结束行，不传读取到文件结束
func (f *File) ReadToSlice(sheet string, startRow int, slice interface{}, endRow ...int) error {
	if reflect.TypeOf(slice).Kind() != reflect.Ptr {
		return errors.New("invalid value, should be pointer to slice")
	}
	value := reflect.ValueOf(slice).Elem() // slice
	itemType := value.Type().Elem()        // item type

	var values = make([]reflect.Value, 0)

	rows, err := f.File.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	var rowIndex = 1
	for rows.Next() {
		if rowIndex < startRow {
			rowIndex++
			continue
		}
		if len(endRow) > 0 && rowIndex > endRow[0] {
			break
		}
		row, err := rows.Columns()
		if err != nil {
			return err
		}

		vE := reflect.New(itemType).Elem()
		item := vE
		if itemType.Kind() == reflect.Ptr && vE.IsNil() && vE.CanSet() {
			// 空指针初始化
			vE.Set(reflect.New(itemType.Elem()))
			item = vE.Elem()
		}
		// 判断是否空行
		isEmptyLine := true
		for i, colCell := range row {
			if i > item.NumField()-1 {
				break
			}
			if len(colCell) > 0 {
				isEmptyLine = false
			}
		}
		if isEmptyLine { // 忽略空行
			continue
		}
		for i, colCell := range row {
			if i > item.NumField()-1 {
				break
			}
			axis := GetAxis(i+1, rowIndex)
			formula, err := f.GetCellFormula(sheet, axis)
			if err != nil {
				return err
			}
			// 公式的值
			var formulaValue string
			if len(formula) > 0 && len(colCell) == 0 {
				// TODO 目前遇到的函数VLOOKUP, 会出现死循环现象
				formulaValue, err = f.CalcCellValue(sheet, axis)
				if err == nil {
					colCell = formulaValue
				}
			}
			switch item.Field(i).Kind() {
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				v, err := cast.ToInt64E(colCell)
				if err != nil {
					return err
				}
				item.Field(i).SetInt(v)
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				v, err := cast.ToUint64E(colCell)
				if err != nil {
					return err
				}
				item.Field(i).SetUint(v)
			case reflect.Float32, reflect.Float64:
				v, err := cast.ToFloat64E(colCell)
				if err != nil {
					return err
				}
				item.Field(i).SetFloat(v)
			case reflect.String:
				item.Field(i).SetString(colCell)
			default:
				return fmt.Errorf("unsupported types: %v", vE.Field(i).Kind())
			}
		}
		values = append(values, vE)
		rowIndex++
	}
	if len(values) > 0 {
		value.Set(reflect.Append(value, values...))
	}
	return nil
}

type rangeInfo struct {
	CopyRow int
	Start   int
	End     int
	CopyNum int
}

func (f *File) beforeTemplateRender(sheet string, data interface{}) error {
	rows, err := f.File.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	rangeStartRg := regexp.MustCompile(`{{\s?range \.(.*?)\s?}}`)
	rangeEndRg := regexp.MustCompile(`{{\s?end\s?}}`)
	rowIndex := 1
	ranges := make([]rangeInfo, 0)
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return err
		}
		for _, colCell := range row {
			if rangeStartRg.MatchString(colCell) {
				// 获取名字
				sub := rangeStartRg.FindStringSubmatch(colCell)
				if len(sub) < 2 {
					return fmt.Errorf("range syntax error")
				}
				v := reflect.ValueOf(data)
				if v.Kind() == reflect.Ptr {
					v = reflect.ValueOf(data).Elem()
				}
				var array = reflect.Value{}
				var temp = v
				for _, n := range strings.Split(sub[1], ".") {
					array = temp.FieldByName(strings.TrimSpace(n))
					temp = array
				}
				if array.Kind() != reflect.Slice && array.Kind() != reflect.Array {
					return fmt.Errorf("%s non-array", sub[1])
				}
				ranges = append(ranges, rangeInfo{
					CopyRow: rowIndex + 1,
					Start:   rowIndex,
					End:     0,
					CopyNum: array.Len(),
				})
			}
			if rangeEndRg.MatchString(colCell) && len(ranges) > 0 {
				ranges[len(ranges)-1].End = rowIndex
			}
		}
		rowIndex++
	}
	if len(ranges) == 0 {
		return nil
	}
	for i := len(ranges) - 1; i >= 0; i-- {
		if ranges[i].CopyNum == 0 {
			for _, r := range []int{ranges[i].End, ranges[i].End - 1, ranges[i].Start} {
				if err := f.RemoveRow(sheet, r); err != nil {
					return err
				}
			}
		} else {
			for j := 0; j < ranges[i].CopyNum-1; j++ {
				if err := f.DuplicateRowTo(sheet, ranges[i].CopyRow, ranges[i].CopyRow+1); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (f *File) GetFirstSheet() string {
	l := f.GetSheetList()
	if len(l) > 0 {
		return l[0]
	}
	return ""
}

// GetLastEmptyRow 获取最后非空行
func (f *File) GetLastEmptyRow(sheet string) (int, error) {
	rows, err := f.File.Rows(sheet)
	if err != nil {
		return 0, err
	}
	defer rows.Close()
	var lastRow = 1
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return 0, err
		}
		l := 0
		for _, colCell := range row {
			l += len(colCell)
		}
		if l == 0 {
			break
		}
		lastRow++
	}
	return lastRow, nil
}

func (f *File) setCell(sheet string, axis string, cell *excelize.Cell) error {
	if err := f.File.SetCellValue(sheet, axis, cell.Value); err != nil {
		return err
	}
	if err := f.File.SetCellStyle(sheet, axis, axis, cell.StyleID); err != nil {
		return err
	}
	if err := f.File.SetCellFormula(sheet, axis, cell.Formula); err != nil {
		return err
	}
	return nil
}

func (s *StreamWriter) SetRowValue(row int, slice []interface{}) error {
	return s.StreamWriter.SetRow(GetAxis(1, row), slice)
}

func (s *StreamWriter) SetRowsValueByTableHeader(startRow int, tableHeader []string, slice interface{}) error {
	valueOf := reflect.ValueOf(slice)
	for i := 0; i < valueOf.Len(); i++ {
		row, err := s.structToRow(valueOf.Index(i), tableHeader)
		if err != nil {
			return nil
		}
		if err := s.SetRowValue(startRow, row); err != nil {
			return err
		}
		startRow++
	}
	return nil
}

func (s *StreamWriter) structToRow(data reflect.Value, tableHeader []string) ([]interface{}, error) {
	var (
		typeOf  = data.Type()
		valueOf = data
	)
	if typeOf.Kind() == reflect.Ptr {
		typeOf = typeOf.Elem()
		valueOf = valueOf.Elem()
	}
	index := make(map[string]int)
	for i, h := range tableHeader {
		index[h] = i
	}
	res := make([]interface{}, len(tableHeader))
	for i := 0; i < typeOf.NumField(); i++ {
		tag, err := ParseTagSetting(typeOf.Field(i), i)
		if err != nil {
			return nil, err
		}
		if tag.Ignore {
			continue
		}
		c, ok := index[tag.Column]
		if !ok {
			continue
		}
		res[c] = valueOf.Field(i).Interface()
	}
	return res, nil
}
