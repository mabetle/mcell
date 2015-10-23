package xlsxsdb

import (
	"github.com/mabetle/mcell/wxlsx"
	"github.com/mabetle/mcore/msdb"
	"github.com/mabetle/mlog"
)

var (
	logger = mlog.GetLogger("github.com/mabetle/mcell/xlsxsdb")
)

// XlsxTable implements msdb.SimpleTable
type XlsxTable struct {
	*msdb.BaseTable
	sheet *wxlsx.Sheet
}

// CheckSimpleTableImpl check
func CheckSimpleTableImpl(file string, sheetName string) (msdb.SimpleTable, error) {
	return NewXlsxTableBySheetName(file, sheetName)
}

// NewSimpleTableBySheetName returns XlsxTable
func NewSimpleTableBySheetName(file string, sheetName string) (*XlsxTable, error) {
	return NewXlsxTableBySheetName(file, sheetName)
}

// NewSimpleTable returns XlsxTable
func NewSimpleTable(file string, sheetIndex int) (*XlsxTable, error) {
	return NewXlsxTable(file, sheetIndex)
}

// NewXlsxTable returns XlsxTable by sheet index.
func NewXlsxTable(file string, sheetIndex int) (*XlsxTable, error) {
	sheet, err := wxlsx.GetSheet(file, sheetIndex)
	if err != nil {
		return nil, err
	}
	return NewXlsxTableBySheet(sheet)
}

// NewXlsxTableBySheetName returns XlsxTable by sheet name.
func NewXlsxTableBySheetName(file string, sheetName string) (*XlsxTable, error) {
	sheet, err := wxlsx.GetSheetByName(file, sheetName)
	if err != nil {
		return nil, err
	}
	return NewXlsxTableBySheet(sheet)
}

// NewXlsxTableBySheet returns XlsxTable
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

// GetHeader returns sheet header row
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

// GetString return colIndex string value
func (t *XlsxTable) GetString(colIndex int) string {
	rowIndex := t.Cusor.RowIndex
	return t.GetRowColString(rowIndex, colIndex)
}

// GetRowColString Random Access
func (t *XlsxTable) GetRowColString(row, col int) string {
	// row or col exceed range.
	if row > t.GetRows() || col > t.GetCols() {
		return ""
	}
	return t.sheet.Rows[row].Cells[col].String()
}
