package wxlsx

import (
	"math"
	"strconv"
	"strings"

	"github.com/mabetle/mcore"
)

// GetLetterIndex returns Excel Cell ColumnIndex
// index start from 0
func GetLetterIndex(letter string) (r int) {
	letter = strings.ToUpper(letter)
	b := []byte(letter)
	for i := 0; i < len(b); i++ {
		seq := int(b[i]) - 64

		if i == len(b)-1 {
			r = seq + r
		} else {
			bNum := seq * int(math.Pow(26, float64(len(b)-i-1)))
			r = bNum + r
		}
	}
	r = r - 1
	return
}

// GetRowColIndex returns Cell row and column index.
// the index start from 0
func GetRowColIndex(cell string) (row, col int) {
	num := "1234567890"
	chars := []byte(cell)
	for index, v := range chars {
		if strings.ContainsAny(num, string(v)) {
			colP := string(chars[:index])
			rowP := string(chars[index:])
			col = GetLetterIndex(colP)
			rowN, err := strconv.ParseInt(rowP, 10, 32)
			if err != nil {
				return
			}
			row = int(rowN) - 1
			return
		}
	}
	return
}

// GetSheetNames returns include sheet names.
func GetSheetNames(location string) ([]string, error) {
	names := []string{}
	wb, err := OpenBook(location)
	if logger.CheckError(err) {
		return names, err
	}
	for _, sheet := range wb.Sheets {
		names = append(names, sheet.Name)
	}
	return names, nil
}

// GetSheetNameByIndex returns index sheet name.
func GetSheetNameByIndex(file string, index int) (string, error) {
	sheet, err := GetSheetByIndex(file, index)
	if logger.CheckError(err) {
		return "", err
	}
	return sheet.Name, nil
}

// GetSheetColumnNames returns Sheet column names, call sheet GetHeaderRowValues
func GetSheetColumnNames(file string, sheetName string) ([]string, error) {
	names := []string{}
	sheet, err := GetSheetByName(file, sheetName)
	if logger.CheckError(err) {
		return names, err
	}
	names = sheet.GetHeaderRowValues()
	return names, nil
}

// IsHasSheet check sheet exists.
func IsHasSheet(file, sheetName string) bool {
	ns, err := GetSheetNames(file)
	if err != nil {
		return false
	}
	return mcore.NewString(sheetName).IsInArrayIgnoreCase(ns)
}
