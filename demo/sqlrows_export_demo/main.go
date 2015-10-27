package main

import (
	"mabetle/libs/hubs"
	"github.com/mabetle/mcell/wxlsx"
	"github.com/mabetle/mlog"
)

var (
	logger = mlog.GetLogger("main")
	sql    = hubs.GetCommonSql()
)

func main() {
	q := "select * from user_info"
	rows, _ := sql.Query(q)

	f, err := wxlsx.SqlRowsToExcel("", rows, "UserName,RealName", "")

	if logger.CheckError(err) {
		return
	}

	f.Save("/rundata/output/demo.xlsx")
}
