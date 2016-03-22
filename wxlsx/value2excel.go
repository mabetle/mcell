package wxlsx

import (
	"fmt"
	"github.com/mabetle/mcore"
	"github.com/mabetle/mmsg"
	"github.com/tealeg/xlsx"
	"reflect"
)

// ValueToExcel value to excel
func ValueToExcel(sheetName string,
	rows interface{},
	include string,
	exclude string) (*xlsx.File, error) {
	// no  locale
	return ValueToExcelWithLocale(sheetName, rows, include, exclude, "", false)
}

// ValueToExcelWithLocale value to excel with locale
// rows should be slice type.
func ValueToExcelWithLocale(sheetName string,
	rows interface{},
	include string,
	exclude string,
	locale string,
	enableLocale bool) (*xlsx.File, error) {

	// check args
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	v := reflect.ValueOf(rows)
	if v.Kind() != reflect.Slice {
		return nil, fmt.Errorf("rows should be a slice")
	}

	file := xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return nil, err
	}
	//fs := mcore.GetArrayFields(rows)
	fs := mcore.GetUsedArrayFields(rows, include, exclude)
	// add header
	row := sheet.AddRow()

	// f means fields
	for _, f := range fs {
		cell := row.AddCell()
		// locale header
		if enableLocale && locale != "" {
			f = mmsg.GetModelFieldLabel(locale, mcore.GetArrayFirstElement(rows), f)
		}
		cell.Value = f
	}

	// add data
	for i := 0; i < v.Len(); i++ {
		rowValue := v.Index(i).Interface()
		row := sheet.AddRow()
		for _, f := range fs {
			cell := row.AddCell()
			// TODO adapter for all kinds data type
			//cell.Value = mcore.GetFieldValue(rowValue, f)
			cv := mcore.GetFieldValueOrigin(rowValue, f)
			cell.SetValue(cv)
		}
	}
	return file, nil
}
