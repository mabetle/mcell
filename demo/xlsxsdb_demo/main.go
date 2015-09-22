package main

import (
	"github.com/mabetle/mcell/xlsxsdb"
	"github.com/mabetle/mcore/msdb"
	"github.com/mabetle/mcore"
	"github.com/mabetle/mcell/wxlsx"
	"github.com/mabetle/mlog"
)

var(
	file="../data/demo.xlsx"
	sheetIndex = 0
)

func PrintSheet(){
	wxlsx.PrintSheetByIndex(file,sheetIndex)
}

func DemoRead(){
	table, err:=xlsxsdb.NewSimpleTable(file, sheetIndex)
	mcore.CheckError(err)
	msdb.DemoSimpleTable(table)
}

func main() {
	//mlog.SetDebugLevel()
	mlog.SetInfoLevel()
	PrintSheet()
	DemoRead()
}


