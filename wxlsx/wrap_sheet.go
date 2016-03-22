package wxlsx

import (
	"github.com/mabetle/mcore"
	"github.com/tealeg/xlsx"
)

// Sheet define sheet
type Sheet struct {
	*xlsx.Sheet
}

// NewSheet create Sheet
func NewSheet(sheet *xlsx.Sheet) *Sheet {
	return &Sheet{Sheet: sheet}
}

// GetCellValue returns cell value
func (s *Sheet) GetCellValue(cell string) (value string) {
	row, col := GetRowColIndex(cell)
	value = s.GetRowColValue(row, col, "")
	return
}

// GetRowColValue returns row col value
// process out of index error
func (s *Sheet) GetRowColValue(row, col int, errDefault string) (value string) {
	// process error
	// index out of range
	defer func() {
		if err := recover(); err != nil {
			logger.Tracef("Error: %v", err)
			value = errDefault
		}
	}()

	rows := s.Rows

	// invalid row and col
	if row < 0 || col < 0 {
		logger.Tracef("Invalid row or col index: RowIndex %d ColIndex %d .", row, col)
		return errDefault
	}

	if row > len(rows) {
		logger.Tracef("row %d exceed range rows %d .", row, len(rows))
		return errDefault
	}
	cells := s.Rows[row].Cells
	if col > len(cells) {
		logger.Tracef("col %d exceed range columns %d .", col, len(cells))
		return errDefault
	}

	value = cells[col].Value
	return
}

// GetHeaderRowValues return header row values
func (s *Sheet) GetHeaderRowValues() (vs []string) {
	if s.MaxRow < 1 {
		// no header row
		return
	}

	for _, cell := range s.Rows[0].Cells {
		cv := cell.Value
		vs = append(vs, cv)
	}

	return
}

// GetColNameIndex returns colname index
// -1 means not found.
func (s *Sheet) GetColNameIndex(colName string) int {
	names := s.GetHeaderRowValues()
	for i, name := range names {
		if mcore.NewString(colName).IsEqualIgnoreCase(name) {
			return i
		}
	}
	return -1
}

// GetCellValueByRowIndexColName returns cell value
func (s *Sheet) GetCellValueByRowIndexColName(rowIndex int, colName string) string {
	colIndex := s.GetColNameIndex(colName)
	return s.GetRowColValue(rowIndex, colIndex, "")
}

// GetCellFloat64ByRowIndexColName return cell float value
func (s *Sheet) GetCellFloat64ByRowIndexColName(rowIndex int, colName string) float64 {
	v := s.GetCellValueByRowIndexColName(rowIndex, colName)
	return mcore.NewString(v).ToFloat64NoError()
}

// GetCellIntByRowIndexColName return cell int value
func (s *Sheet) GetCellIntByRowIndexColName(rowIndex int, colName string) int {
	v := s.GetCellValueByRowIndexColName(rowIndex, colName)
	return mcore.NewString(v).ToIntNoError()
}
