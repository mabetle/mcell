package wxlsx

import (
	"database/sql"
	"github.com/tealeg/xlsx"
	"github.com/github.com/mabetle/mcore"
	"github.com/mabetle/mmsg"
)


// SqlRowsToExcel
// disable locale
func SqlRowsToExcel(sheetName string, 
	rows *sql.Rows, 
	include string, 
	exclude string) (*xlsx.File, error) {
	
		return SqlRowsToExcelWithLocale(sheetName,"", rows,include,exclude,"",false)

}

// SqlRowsToLocalHeaderExcel
// Locale message, table column name.
// params:
//	sheetName
//	rows
//	include
//	exclude
//	locale 
//	enableLocale
func SqlRowsToExcelWithLocale(sheetName string, 
	tableName string,
	rows *sql.Rows, 
	include string, 
	exclude string,
	locale string,
	enableLocale bool) (*xlsx.File, error) {

		
	defer rows.Close()

	if sheetName == "" {
		sheetName = "Sheet1"
	}
	
	file := xlsx.NewFile()
	sheet := file.AddSheet(sheetName)

	colNames, err := rows.Columns()
	if err != nil {
		return nil, err
	}

	// add header
	row := sheet.AddRow()
	for _, colName := range colNames {
		if !mcore.IsIncludeExcludeIn(colName, colNames, include, exclude) {
			continue
		}
		cell := row.AddCell()
		// colName to locale label
		
		if enableLocale && locale != "" {
			colName = mmsg.GetTableColumnLabel(locale, tableName,colName)
		}

		cell.Value = colName
	}

	scanArgs := make([]interface{}, len(colNames))
	values := make([]interface{}, len(colNames))

	for i := range values {
		scanArgs[i] = &values[i]
	}

	// add rows data
	for rows.Next() {
		err := rows.Scan(scanArgs...)
		if err != nil {
			continue
		}

		row := sheet.AddRow()

		index := -1
		for _, v := range values {
			index++
			// skip for no include column
			if !mcore.IsIncludeExcludeIn(colNames[index], colNames, include, exclude) {
				continue
			}
			cell := row.AddCell()
			cell.SetValue(v)
		}
	}

	return file, nil
}
