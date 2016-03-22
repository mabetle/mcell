package wxlsx

import (
	"encoding/json"
	"fmt"
	"github.com/mabetle/mcore"
	"github.com/mabetle/mmsg"
	"github.com/tealeg/xlsx"
)

/*
	go get -u -v github.com/tealeg/xlsx
*/

// ArrayToExcel write data
func ArrayToExcel(sheetName string, data [][]string) (*xlsx.File, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	file := xlsx.NewFile()

	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return nil, err
	}
	for _, rowData := range data {
		row := sheet.AddRow()
		for _, colData := range rowData {
			cell := row.AddCell()
			cell.Value = colData
		}
	}

	return file, nil
}

// GetMapKeys return map keys
func GetMapKeys(m map[string]interface{}) (keys []string) {
	for k := range m {
		keys = append(keys, k)
	}
	return
}

// JSONToExcel define values
// jsData should contain a array
func JSONToExcel(
	sheetName string,
	jsData []byte,
	include string,
	exclude string,
	locale string,
) (*xlsx.File, error) {
	var rows []map[string]interface{}
	err := json.Unmarshal(jsData, &rows)
	if err != nil {
		return nil, err
	}
	if len(rows) < 1 {
		return nil, fmt.Errorf("No datas found")
	}
	headMap := rows[0]
	allKeys := GetMapKeys(headMap)
	keys := mcore.GetFieldsUsed(allKeys, include, exclude)

	if sheetName == "" {
		sheetName = "Sheet1"
	}
	file := xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return nil, err
	}
	// add header row
	row := sheet.AddRow()

	for _, key := range keys {
		cell := row.AddCell()
		cell.Value = mmsg.GetTableColumnLabel(locale, "", key)
	}

	// add datas
	for _, row := range rows {
		sheetRow := sheet.AddRow()
		for _, key := range keys {
			cell := sheetRow.AddCell()
			value := row[key]
			cell.SetValue(value)
		}
	}
	return file, nil
}
