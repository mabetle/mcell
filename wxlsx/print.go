package wxlsx

import (
	"fmt"

	"github.com/mabetle/mcore"
)

// PrintSheet utils
func PrintSheet(sheet *Sheet) {
	fmt.Println("====Begin Print Sheet====")
	if nil == sheet {
		fmt.Println("Sheet is nil")
	}

	for _, row := range sheet.Rows {
		//rows
		for _, cell := range row.Cells {
			//col
			cellStr, _ := cell.String()
			fmt.Printf("%s \t", cellStr)
		}
		fmt.Println()
	}

	fmt.Println("====End.. Print Sheet====")
}

// Print print
func (sheet *Sheet) Print() {
	PrintSheet(sheet)
}

// PrintCell cell format example: A1
func (sheet *Sheet) PrintCell(cell string) {
	fmt.Printf("Cell: %s, Value: %s\n", cell, sheet.GetCellValue(cell))
}

// PrintSheetByIndex print
func PrintSheetByIndex(file string, sheetIndex int) {
	sheet, err := GetSheetByIndex(file, sheetIndex)
	if err != nil {
		fmt.Println(err)
		return
	}
	PrintSheet(sheet)
}

// PrintSheetByName print
func PrintSheetByName(file string, sheetName string) {
	sheet, err := GetSheetByName(file, sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	PrintSheet(sheet)
}

// PrintSheetNames print
func PrintSheetNames(file string) {
	vs, _ := GetSheetNames(file)
	mcore.PrintStringArray(vs)
}
