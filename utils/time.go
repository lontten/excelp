package utils

import (
	"reflect"
	"time"

	"github.com/lontten/lcore/v2/types"
)

var timeMap = map[reflect.Type]bool{
	reflect.TypeOf(types.LocalDate{}):     true,
	reflect.TypeOf(types.LocalTime{}):     true,
	reflect.TypeOf(types.LocalDateTime{}): true,
	reflect.TypeOf(time.Time{}):           true,
}

func IsTimeType(v reflect.Value) bool {
	_, ok := timeMap[v.Type()]
	if ok {
		return true
	}
	return false
}
