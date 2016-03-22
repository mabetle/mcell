package wxlsx

import (
	"github.com/tealeg/xlsx"
)

// OpenBook returns Workbook
func OpenBook(location string) (*xlsx.File, error) {
	return xlsx.OpenFile(location)
}

// GetRowColValue returns excel row col value
// parameters: location sheetname, row,col
func GetRowColValue(location, sheetName string, row, col int, errDefault string) string {
	sheet, err := GetSheetByName(location, sheetName)
	// error
	if err != nil {
		logger.Tracef("GetRowColValue: \n\tFile: %s\n\tSheetName: %s\n\tRow: %d Col: %d\n\tError:%v",
			location,
			sheetName,
			row, col,
			err)
		return errDefault
	}
	return sheet.GetRowColValue(row, col, errDefault)
}

// GetSheetRowColValue returns sheet row col value
func GetSheetRowColValue(sheet *Sheet, row, col int, errDefault string) string {
	return sheet.GetRowColValue(row, col, errDefault)
}

// GetCellValue return file sheet cell value
// cell is Excel format. eg: AA23
func GetCellValue(location, sheetName, cell, errDefault string) string {
	row, col := GetRowColIndex(cell)
	return GetRowColValue(location, sheetName, row, col, errDefault)
}

// GetSheetCellValue returns sheet cell value
func GetSheetCellValue(sheet *Sheet, cell, errDefault string) string {
	row, col := GetRowColIndex(cell)
	return GetSheetRowColValue(sheet, row, col, errDefault)
}

// GetCellsValues returns cells values
func GetCellsValues(location, sheetName string, cells []string, errDefault string) (values []string) {
	for _, cell := range cells {
		v := GetCellValue(location, sheetName, cell, errDefault)
		values = append(values, v)
	}
	return
}

// GetSheetCellsValues return cells values.
func GetSheetCellsValues(sheet *Sheet, cells []string, errDefault string) (values []string) {
	for _, cell := range cells {
		v := GetSheetCellValue(sheet, cell, errDefault)
		values = append(values, v)
	}
	return
}
