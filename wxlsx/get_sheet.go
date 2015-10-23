package wxlsx

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// same as GetSheetByIndex
func GetSheet(file string, sheetIndex int) (*Sheet, error) {
	return GetSheetByIndex(file, sheetIndex)
}

// GetSheetByIndex
// index from 0
func GetSheetByIndex(file string, sheetIndex int) (*Sheet, error) {
	book, err := xlsx.OpenFile(file)
	if nil != err {
		return nil, err
	}
	sheetNums := len(book.Sheets)
	if sheetIndex > sheetNums {
		return nil, fmt.Errorf("Open %s index %d out of sheet index, max sheet index is %d", file, sheetIndex, sheetNums)
	}
	return NewSheet(book.Sheets[sheetIndex]), nil
}

// GetSheetByName
func GetSheetByName(file string, sheetName string) (*Sheet, error) {
	book, err := xlsx.OpenFile(file)
	if nil != err {
		return nil, err
	}

	sheetNums := len(book.Sheets)
	for sheetIndex := 0; sheetIndex < sheetNums; sheetIndex++ {
		curSheet := book.Sheets[sheetIndex]
		curName := curSheet.Name
		if curName == sheetName {
			return NewSheet(curSheet), nil
		}
	}
	//not found sheet
	return nil, fmt.Errorf("SheetName %s not found in %s.", sheetName, file)
}
