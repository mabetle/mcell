package xlsxsdb

import (
	"fmt"
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

// NewSimpleTableFromByteArray creates XlsxTable
func NewSimpleTableFromByteArray(data []byte, sheetIndex int) (*XlsxTable, error) {
	sheet, err := wxlsx.GetSheetFromByteArray(data, sheetIndex)
	if err != nil {
		return nil, err
	}
	return NewXlsxTableBySheet(sheet)
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
		colName, _ := cells[col].String()
		colNames = append(colNames, colName)
	}
	//	sheet.Rows[0]
	return colNames
}

// GetString return colIndex string value
func (t *XlsxTable) GetStringByIndex(colIndex int) string {
	rowIndex := t.Cusor.RowIndex
	return t.GetRowColString(rowIndex, colIndex)
}

// GetRowColString Random Access
func (t *XlsxTable) GetRowColString(row, col int) string {
	defer func() {
		if err := recover(); err != nil {
			logger.Warnf("Error GetRowColValue, Row: %d, MaxRow: %d, Col: %d, MaxCol: %d, Error:%v",
				row,
				t.GetRows(),
				col,
				t.GetCols(),
				err)
			//	return ""
		}
	}()
	// row or col exceed range.
	if row > t.GetRows() || col > t.GetCols() {
		return ""
	}
	// value as original
	return t.sheet.Rows[row].Cells[col].Value
}

func (t *XlsxTable) Print() {
	cols := t.GetColNames()
	row := 0
	for t.Next() {
		row++
		fmt.Printf("\nRow: %d\n", row)
		for _, col := range cols {
			fmt.Printf("%s:%s\n", col, t.GetString(col))
		}
	}
}
