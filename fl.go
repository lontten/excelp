package excelp

import (
	"github.com/lontten/excelp/utils"
	"github.com/pkg/errors"
	"reflect"
	"strconv"
	"strings"
	"unicode"
)

// excel row index:字段名
// row index 从0开始
// 过滤掉首字母小写的字段
func _getStructC(t reflect.Type) (m map[int]Field, err error) {
	m = make(map[int]Field)
	numField := t.NumField()
	for i := 0; i < numField; i++ {
		structField := t.Field(i)
		if structField.Anonymous {
			data, err := _getStructC(structField.Type)
			if err != nil {
				return nil, err
			}
			for k, v := range data {
				m[k] = v
			}
			continue
		}

		name := structField.Name

		// 过滤掉首字母小写的字段
		if unicode.IsLower([]rune(name)[0]) {
			continue
		}

		tag := structField.Tag.Get("excelp")
		index, err := getIndex(tag)
		if err != nil {
			return nil, err
		}
		if index == -1 {
			continue
		}
		fieldConfig := Field{
			name:     name,
			required: getRequired(tag),
		}
		m[index] = fieldConfig
	}
	return m, nil
}

// 没有设置index，返回-1
func getIndex(s string) (int, error) {
	split := strings.Split(s, ";")
	for _, s2 := range split {
		index := strings.Index(s2, ":")
		if index > -1 {
			head := s2[:index]
			if head == "index" {
				end := s2[index+1:]
				atoi, err := strconv.Atoi(end)
				if err == nil {
					return atoi, nil
				}
				number, err := utils.ColumnNameToNumber(end)
				if err == nil {
					return number, nil
				}
				return -1, errors.New("excelp index error")
			}
		}
	}
	return -1, nil
}

func getTimeFormat(s string) string {
	split := strings.Split(s, ";")
	for _, s2 := range split {
		index := strings.Index(s2, ":")
		if index > -1 {
			head := s2[:index]
			if head == "format" {
				return s2[index+1:]
			}
		}
	}
	return ""
}

func getRequired(s string) bool {
	split := strings.Split(s, ";")
	for _, s2 := range split {
		if s2 == "required" {
			return true
		}
	}
	return false
}
