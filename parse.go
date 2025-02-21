package excelp

import (
	"database/sql"
	"fmt"
	"github.com/lontten/excelp/utils"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
	"strings"
)

func parse[T any](c *ExcelReadContext, index int, row []string) (T, []CellErr) {
	rowLen := len(row)
	dest := new(T)
	errList := make([]CellErr, 0)
	v := reflect.ValueOf(dest).Elem()
	for i, s := range c.m {
		if i >= rowLen {
			continue
		}
		value := row[i]
		fieldByName := v.FieldByName(s.name)
		err := scanField(fieldByName, value, s)
		if err != nil {
			name, _ := utils.ColumnNumberToName(i)
			errList = append(errList, CellErr{
				Err:   err.Error(),
				Col:   name,
				Row:   index,
				Value: value,
			})
		}
	}
	return *dest, errList
}

func scanField(field reflect.Value, value string, f Field) error {
	if value == "" {
		if f.required {
			return errors.New("required")
		}
		return nil
	}
	// 如果字段是指针类型
	if field.Kind() == reflect.Ptr {
		// 如果指针为 nil，需要分配内存
		if field.IsNil() {
			field.Set(reflect.New(field.Type().Elem())) // 分配内存
		}

		// 获取指针指向的值
		field = field.Elem()
	}

	// 尝试调用字段的 Scan 方法
	if scanner, ok := field.Addr().Interface().(sql.Scanner); ok {
		var src any = value

		if field.Kind() == reflect.Slice && field.Type().Elem().Kind() == reflect.Uint8 {
			src = []byte(value)
		} else {
			if utils.IsTimeType(field) {
				timeFloat, err := strconv.ParseFloat(value, 64)
				if err != nil {
					return fmt.Errorf("can not convert %v to time.Time", value)
				}
				t, err := excelize.ExcelDateToTime(timeFloat, false)
				if err != nil {
					return fmt.Errorf("can not convert %v to time.Time", value)
				}
				src = t
			}
		}
		if err := scanner.Scan(src); err != nil {
			return fmt.Errorf("scan failed for field %s: %v", field.Type().Name(), err)
		}
		return nil
	}

	// 如果字段没有实现 sql.Scanner，尝试手动处理基本类型
	switch field.Kind() {
	case reflect.String:
		field.SetString(value)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		value = strings.ReplaceAll(value, ",", "")
		intValue, err := strconv.ParseInt(value, 10, 64)
		if err != nil {
			return fmt.Errorf("failed to convert %s to int for field %s: %v", value, field.Type().Name(), err)
		}
		field.SetInt(intValue)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		value = strings.ReplaceAll(value, ",", "")
		uintValue, err := strconv.ParseUint(value, 10, 64)
		if err != nil {
			return fmt.Errorf("failed to convert %s to uint for field %s: %v", value, field.Type().Name(), err)
		}
		field.SetUint(uintValue)
	case reflect.Float32, reflect.Float64:
		value = strings.ReplaceAll(value, ",", "")
		floatValue, err := strconv.ParseFloat(value, 64)
		if err != nil {
			return fmt.Errorf("failed to convert %s to float for field %s: %v", value, field.Type().Name(), err)
		}
		field.SetFloat(floatValue)
	case reflect.Bool:
		boolValue, err := strconv.ParseBool(value)
		if err != nil {
			return fmt.Errorf("failed to convert %s to bool for field %s: %v", value, field.Type().Name(), err)
		}
		field.SetBool(boolValue)
	case reflect.Struct:
		switch field.Type().Name() {
		case "Time":
			timeFloat, err := strconv.ParseFloat(value, 64)
			if err != nil {
				return fmt.Errorf("can not convert %v to time.Time", value)
			}
			t, err := excelize.ExcelDateToTime(timeFloat, false)
			if err != nil {
				return fmt.Errorf("can not convert %v to time.Time", value)
			}
			field.Set(reflect.ValueOf(t))
		}
	default:
		fmt.Println("field:", field.Type().Name(), "kind:", field.Kind().String())
		return fmt.Errorf("unsupported field type: %s, kind: %s", field.Type().Name(), field.Kind().String())
	}

	return nil
}
