package main

import (
	"mabetle/libs/dbc"
	"mabetle/libs/wage_record/models"
	"github.com/mabetle/mcell/wxlsx"
	"github.com/mabetle/mlog"
)

var (
	logger = mlog.GetLogger("main")
	xorm   = dbc.GetXorm()
)

func main() {
	m := &models.WageRecord{}
	m.UserName = "zsc"
	m.TheYear = 2014
	var rows []models.WageRecord
	xorm.Find(&rows, m)
	f, err := wxlsx.ValueToExcel("", rows, "", "Id")
	if logger.CheckError(err) {
		return
	}
	f.Save("/rundata/output/demo.xlsx")
}
