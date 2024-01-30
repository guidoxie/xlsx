package xlsx

import (
	"encoding/json"
	"fmt"
	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
	"reflect"
	"regexp"
	"strings"
	"sync"
)

type TagSetting struct {
	Axis     string          // 写入坐标
	Style    *excelize.Style // 单元格样式
	Column   string          // 字段名
	ColWidth float64         // 列宽度
	Col      string          // 列
	Ignore   bool            // 忽略字段
}

const (
	tagXlsx     = "xlsx"
	tagAxis     = "axis"
	tagStyle    = "style"
	tagColumn   = "column"
	tagColWidth = "colWidth"
	sep         = ";"
)

var parseTagSettingCache = sync.Map{}

func ParseTagSetting(sf reflect.StructField, field int, row ...int) (*TagSetting, error) {
	tag := sf.Tag.Get(tagXlsx)
	key := fmt.Sprintf("%s.%d", tag, field)
	if len(row) > 0 {
		key = fmt.Sprintf("%s.%d", key, row[0])
	}
	if ts, ok := parseTagSettingCache.Load(key); ok {
		return ts.(*TagSetting), nil
	}
	settings := map[string]string{}
	if tag == "-" || tag == "" {
		return &TagSetting{Ignore: true}, nil
	}
	names := strings.Split(tag, sep)
	for i := 0; i < len(names); i++ {
		j := i
		if len(names[j]) > 0 {
			for {
				if names[j][len(names[j])-1] == '\\' {
					i++
					names[j] = names[j][0:len(names[j])-1] + sep + names[i]
					names[i] = ""
				} else {
					break
				}
			}
		}
		values := strings.Split(names[j], ":")
		k := strings.TrimSpace(values[0])

		if len(values) >= 2 {
			settings[k] = strings.Join(values[1:], ":")
		} else if k != "" {
			settings[k] = k
		}
	}
	res := &TagSetting{}
	for k, v := range settings {
		switch k {
		case tagAxis:
			if len(regexp.MustCompile("\\d+").FindStringSubmatch(v)) == 0 && len(row) > 0 {
				v += cast.ToString(row[0])
			}
			res.Axis = v
		case tagStyle:
			res.Style = &excelize.Style{}
			if err := json.Unmarshal([]byte(v), res.Style); err != nil {
				return nil, err
			}
		case tagColumn:
			res.Column = v
		case tagColWidth:
			res.ColWidth = cast.ToFloat64(v)
		}
	}
	if len(res.Axis) == 0 {
		res.Axis = GetAxis(field+1, row...)
	}
	res.Col = regexp.MustCompile("\\d+").ReplaceAllString(res.Axis, "")
	parseTagSettingCache.Store(key, res)
	return res, nil
}

// GetAxis 坐标从1开始
func GetAxis(column int, row ...int) string {
	var (
		str  = ""
		k    int
		temp []int //保存转化后每一位数据的值，然后通过索引的方式匹配A-Z
	)
	//用来匹配的字符A-Z
	slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if column > 26 { //数据大于26需要进行拆分
		for {
			k = column % 26 //从个位开始拆分，如果求余为0，说明末尾为26，也就是Z，如果是转化为26进制数，则末尾是可以为0的，这里必须为A-Z中的一个
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			column = (column - k) / 26 //减去Num最后一位数的值，因为已经记录在temp中
			if column <= 26 {          //小于等于26直接进行匹配，不需要进行数据拆分
				temp = append(temp, column)
				break
			}
		}
	} else {
		if len(row) > 0 {
			return fmt.Sprintf("%s%d", slice[column], row[0])
		}
		return slice[column]
	}

	for _, value := range temp {
		str = slice[value] + str //因为数据切分后存储顺序是反的，所以Str要放在后面
	}
	if len(row) > 0 {
		return fmt.Sprintf("%s%d", str, row[0])
	}
	return str
}
