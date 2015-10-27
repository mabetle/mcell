package main

import (
	"github.com/mabetle/mcell/wxlsx"
	"github.com/mabetle/mlog"
)

var (
	logger = mlog.GetLogger("main")
)

func main() {
	location := "../data/demo.xlsx"
	sheet, err := wxlsx.GetSheetByIndex(location, 0)
	logger.CheckError(err)
	wxlsx.PrintSheet(sheet)
}
