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

// WriteData
func ArrayToExcel(sheetName string, data [][]string) (*xlsx.File, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	file := xlsx.NewFile()

	sheet := file.AddSheet(sheetName)

	for _, rowData := range data {
		row := sheet.AddRow()
		for _, colData := range rowData {
			cell := row.AddCell()
			cell.Value = colData
		}
	}

	return file, nil
}

func GetMapKeys(m map[string]interface{}) (keys []string) {
	for k, _ := range m {
		keys = append(keys, k)
	}
	return
}

// JsonToExcel
// jsData should contain a array
func JsonToExcel(
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
	sheet := file.AddSheet(sheetName)
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
