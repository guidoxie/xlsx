package xlsx

import (
	"fmt"
	"reflect"
	"testing"
)

func TestParseTagSetting(t *testing.T) {
	s1 := struct {
		Name string `json:"name" xlsx:"column:名字"`
		Age  int    `json:"age" xlsx:"column:年龄"`
	}{
		Name: "s1",
		Age:  10,
	}

	s2 := struct {
		Name   string `json:"name" xlsx:"column:名字"`
		Age    int    `json:"age" xlsx:"column:年龄"`
		Height int    `json:"height" xlsx:"column:身高"`
	}{
		Name:   "s2",
		Age:    11,
		Height: 178,
	}
	v1 := reflect.TypeOf(s1)

	for i := 0; i < v1.NumField(); i++ {
		ts, err := ParseTagSetting(v1.Field(i), i)
		if err != nil {
			t.Fatal(err)
		}
		fmt.Println(ts)
	}
	v2 := reflect.TypeOf(s2)
	for i := 0; i < v2.NumField(); i++ {
		ts, err := ParseTagSetting(v2.Field(i), i)
		if err != nil {
			t.Fatal(err)
		}
		fmt.Println(ts)
	}
}
