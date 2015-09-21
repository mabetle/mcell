package wxlsx

import (
	"github.com/tealeg/xlsx"
)

// OpenBook
func OpenBook(location string) (*xlsx.File, error) {
	return xlsx.OpenFile(location)
}

// GetSheetNames
func GetSheetNames(location string) (r []string) {
	book, err := xlsx.OpenFile(location)
	if err != nil {
		return
	}
	for _, v := range book.Sheets {
		r = append(r, v.Name)
	}
	return
}

// GetCellValue
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
	return sheet.GetRowColValue(row,col,errDefault)
}

// GetSheetRowColValue
func GetSheetRowColValue(sheet *Sheet, row, col int, errDefault string) string {
	return sheet.GetRowColValue(row,col,errDefault)
}

// GetCellValue
// cell is Excel format. eg: AA23
func GetCellValue(location, sheetName, cell, errDefault string) string {
	row, col := GetRowColIndex(cell)
	return GetRowColValue(location, sheetName, row, col, errDefault)
}

// GetSheetCellValue
func GetSheetCellValue(sheet *Sheet, cell, errDefault string) string {
	row, col := GetRowColIndex(cell)
	return GetSheetRowColValue(sheet, row, col, errDefault)
}

// GetCellsValues
func GetCellsValues(location, sheetName string, cells []string, errDefault string) (values []string) {
	for _, cell := range cells {
		v := GetCellValue(location, sheetName, cell, errDefault)
		values = append(values, v)
	}
	return
}

// GetSheetCellsValues
func GetSheetCellsValues(sheet *Sheet, cells []string, errDefault string) (values []string) {
	for _, cell := range cells {
		v := GetSheetCellValue(sheet, cell, errDefault)
		values = append(values, v)
	}
	return
}
