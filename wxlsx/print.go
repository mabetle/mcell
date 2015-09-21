package wxlsx

import (
	"fmt"
	"github.com/github.com/mabetle/mcore"
)

// utils
func PrintSheet(sheet *Sheet) {
	fmt.Println("====Begin Print Sheet====")
	if nil == sheet {
		fmt.Println("Sheet is nil")
	}

	for _, row := range sheet.Rows {
		//rows
		for _, cell := range row.Cells {
			//col
			fmt.Printf("%s\t", cell.String())
		}
		fmt.Println()
	}

	fmt.Println("====End.. Print Sheet====")
}

func (sheet *Sheet) Print() {
	PrintSheet(sheet)
}

// Cell format example: A1
func (sheet *Sheet) PrintCell(cell string) {
	fmt.Printf("Cell: %s, Value: %s\n", cell, sheet.GetCellValue(cell))
}

func PrintSheetByIndex(file string, sheetIndex int) {
	sheet, err := GetSheetByIndex(file, sheetIndex)
	if err != nil {
		fmt.Println(err)
		return
	}
	PrintSheet(sheet)
}

func PrintSheetByName(file string, sheetName string) {
	sheet, err := GetSheetByName(file, sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	PrintSheet(sheet)
}

// PrintSheetNames
func PrintSheetNames(file string) {
	vs := GetSheetNames(file)
	mcore.PrintStringArray(vs)
}
