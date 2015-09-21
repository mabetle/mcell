package xlsxsdb

import (
	"fmt"
	"github.com/mabetle/mcell/wxlsx"
	"github.com/github.com/mabetle/mcore/msdb"
	"github.com/mabetle/mlog"
)

var (
	logger = mlog.GetLogger("github.com/mabetle/mcell/xlsxsdb")
)

//implement msdb.SimpleTable
type XlsxTable struct {
	*msdb.BaseTable
	sheet *wxlsx.Sheet
}

func NewSimpleTableBySheetName(file string, sheetName string) (msdb.SimpleTable, error) {
	return NewXlsxTableBySheetName(file, sheetName)
}

func NewSimpleTable(file string, sheetIndex int) (msdb.SimpleTable, error) {
	return NewXlsxTable(file, sheetIndex)
}

func NewXlsxTable(file string, sheetIndex int) (*XlsxTable, error) {
	sheet, err := wxlsx.GetSheet(file, sheetIndex)
	if err != nil {
		return nil, err
	}
	return NewXlsxTableBySheet(sheet)
}

func NewXlsxTableBySheetName(file string, sheetName string) (*XlsxTable, error) {
	sheet, err := wxlsx.GetSheetByName(file, sheetName)
	if err != nil {
		return nil, err
	}
	return NewXlsxTableBySheet(sheet)
}

func NewXlsxTableBySheet(sheet *wxlsx.Sheet) (*XlsxTable, error) {
	table := new(XlsxTable)
	bt := new(msdb.BaseTable)
	cu := new(msdb.Cusor)

	table.sheet = sheet
	cu.MaxIndex = len(sheet.Rows) - 1
	bt.Cusor = cu
	bt.Header = GetHeader(sheet)

	table.BaseTable = bt
	table.StringGetter = table
	return table, nil
}

func GetHeader(sheet *wxlsx.Sheet) []string {
	// define a slice
	var colNames []string
	cells := sheet.Rows[0].Cells
	cols := len(cells)
	for col := 0; col < cols; col++ {
		colName := cells[col].String()
		colNames = append(colNames, colName)
	}
	//	sheet.Rows[0]
	return colNames
}

func (t *XlsxTable) GetString(colIndex int) (result string) {
	defer func() {
		if err := recover(); err != nil {
			logger.Error(err)
			return
		}
	}()
	logger.Trace("Row:", t.Cusor.RowIndex, "Col:", colIndex)
	result = t.sheet.Rows[t.Cusor.RowIndex].Cells[colIndex].String()
	return
}

// Random Access
func (t *XlsxTable) GetRowColString(row, col int) (value string) {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	value = t.sheet.Rows[row].Cells[col].String()
	return
}
