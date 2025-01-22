package excelp

import (
	"reflect"
	"testing"
	"time"
)

func Test_scanField(t *testing.T) {
	type args struct {
		field      reflect.Value
		value      string
		timeFormat string
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{
			name: "int",
			args: args{
				field: reflect.ValueOf(new(int)),
				value: "1",
			},
			wantErr: false,
		},
		{
			name: "int64",
			args: args{
				field: reflect.ValueOf(new(int64)),
				value: "1",
			},
			wantErr: false,
		},
		{
			name: "float64",
			args: args{
				field: reflect.ValueOf(new(float64)),
				value: "1.1",
			},
			wantErr: false,
		},
		{
			name: "string",
			args: args{
				field: reflect.ValueOf(new(string)),
				value: "1",
			},
			wantErr: false,
		},
		{
			name: "date",
			args: args{
				field:      reflect.ValueOf(new(time.Time)),
				value:      "2020-01-01",
				timeFormat: "2006-01-02",
			},
			wantErr: false,
		},
		{
			name: "datetime",
			args: args{
				field:      reflect.ValueOf(new(time.Time)),
				value:      "2020-01-01 01:01:01",
				timeFormat: "2006-01-02 15:04:05",
			},
			wantErr: false,
		},
		{
			name: "time",
			args: args{
				field:      reflect.ValueOf(new(time.Time)),
				value:      "01:01:01",
				timeFormat: "15:04:05",
			},
			wantErr: false,
		},
		{
			name: "error",
			args: args{
				field: reflect.ValueOf(new(int)),
				value: "a",
			},
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err := scanField(tt.args.field, tt.args.value, tt.args.timeFormat); (err != nil) != tt.wantErr {
				t.Errorf("scanField() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}
