package wxlsx

import (
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

// JsonToExcel
//
func JsonToExcel(sheetName string, jsData []byte) (*xlsx.File, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	file := xlsx.NewFile()
	sheet := file.AddSheet(sheetName)

	sheet.AddRow()

	return file, nil

}
