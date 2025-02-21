package utils

import (
	"github.com/lontten/lcore/types"
	"reflect"
	"time"
)

var timeMap = map[reflect.Type]bool{
	reflect.TypeOf(types.Date{}):     true,
	reflect.TypeOf(types.Time{}):     true,
	reflect.TypeOf(types.DateTime{}): true,
	reflect.TypeOf(time.Time{}):      true,
}

func IsTimeType(v reflect.Value) bool {
	_, ok := timeMap[v.Type()]
	if ok {
		return true
	}
	return false
}
