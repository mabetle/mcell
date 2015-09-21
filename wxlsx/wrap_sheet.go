package wxlsx

import (
	"github.com/mabetle/mcore"
	"github.com/tealeg/xlsx"
)

type Sheet struct {
	*xlsx.Sheet
}

// NewSheet
func NewSheet(sheet *xlsx.Sheet) *Sheet {
	return &Sheet{Sheet: sheet}
}

// GetCellValue
func (s *Sheet) GetCellValue(cell string) (value string) {
	row, col := GetRowColIndex(cell)
	value = s.GetRowColValue(row, col, "")
	return
}

// GetRowColValue
// process out of index error
func (sheet *Sheet) GetRowColValue(row, col int, errDefault string) (value string) {
	// process error
	// index out of range
	defer func() {
		if err := recover(); err != nil {
			logger.Tracef("Error: %v", err)
			value = errDefault
		}
	}()

	rows := sheet.Rows

	// invalid row and col
	if row < 0 || col < 0 {
		logger.Tracef("Invalid row or col index: RowIndex %d ColIndex %d .", row, col)
		return errDefault
	}

	if row > len(rows) {
		logger.Tracef("row %d exceed range rows %d .", row, len(rows))
		return errDefault
	}
	cells := sheet.Rows[row].Cells
	if col > len(cells) {
		logger.Tracef("col %d exceed range columns %d .", col, len(cells))
		return errDefault
	}

	value = cells[col].String()
	return
}

// GetHeaderRowValues
func (sheet *Sheet) GetHeaderRowValues() (vs []string) {
	if sheet.MaxRow < 1 {
		// no header row
		return
	}

	for _, cell := range sheet.Rows[0].Cells {
		cv := cell.String()
		vs = append(vs, cv)
	}

	return
}

// GetColNameIndex
// -1 means not found.
func (sheet *Sheet) GetColNameIndex(colName string) int {
	names := sheet.GetHeaderRowValues()
	for i, name := range names {
		if mcore.NewString(colName).IsEqualIgnoreCase(name) {
			return i
		}
	}
	return -1
}

// GetCellValueByRowIndexColName
func (sheet *Sheet) GetCellValueByRowIndexColName(rowIndex int, colName string) string {
	colIndex := sheet.GetColNameIndex(colName)
	return sheet.GetRowColValue(rowIndex, colIndex, "")
}

// GetCellFloat64ByRowIndexColName
func (sheet *Sheet) GetCellFloat64ByRowIndexColName(rowIndex int, colName string) float64 {
	v := sheet.GetCellValueByRowIndexColName(rowIndex, colName)
	return mcore.NewString(v).ToFloat64NoError()
}

// GetCellIntByRowIndexColName
func (sheet *Sheet) GetCellIntByRowIndexColName(rowIndex int, colName string) int {
	v := sheet.GetCellValueByRowIndexColName(rowIndex, colName)
	return mcore.NewString(v).ToIntNoError()
}
