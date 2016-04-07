package wxlsx

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// GetSheet same as GetSheetByIndex
func GetSheet(file string, sheetIndex int) (*Sheet, error) {
	return GetSheetByIndex(file, sheetIndex)
}

// GetSheetFromByteArray returns Sheet
func GetSheetFromByteArray(data []byte, sheetIndex int) (*Sheet, error) {
	book, err := xlsx.OpenBinary(data)
	if err != nil {
		return nil, err
	}
	return GetSheetFromBookByIndex(book, sheetIndex)
}

// GetSheetByIndex returns Sheet
// index from 0
func GetSheetByIndex(file string, sheetIndex int) (*Sheet, error) {
	book, err := xlsx.OpenFile(file)
	if nil != err {
		return nil, err
	}
	return GetSheetFromBookByIndex(book, sheetIndex)
}

// GetSheetFromBookByIndex returns Sheet
func GetSheetFromBookByIndex(book *xlsx.File, sheetIndex int) (*Sheet, error) {
	sheetNums := len(book.Sheets)
	if sheetIndex > sheetNums {
		return nil, fmt.Errorf("index %d out of sheet index, max sheet index is %d", sheetIndex, sheetNums)
	}
	return NewSheet(book.Sheets[sheetIndex]), nil
}

// GetSheetByName returns Sheet
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
