package wxlsx

import (
	"errors"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
	"strings"
)

// ExportCSV export
func ExportCSV(excelFileName string, sheetIndex int, delimiter string, out io.Writer) error {

	xlFile, error := xlsx.OpenFile(excelFileName)
	// check file exists
	if error != nil {
		return error
	}

	sheetLen := len(xlFile.Sheets)
	// check sheet exists.
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, sheetIndex between 0 and %d\n", sheetIndex, sheetLen-1)
	}

	if delimiter == "" {
		delimiter = ","
	}

	sheet := xlFile.Sheets[sheetIndex]
	// walk rows and cols
	for _, row := range sheet.Rows {
		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
				vals = append(vals, fmt.Sprintf("%q", cell.Value))
			}
			line := strings.Join(vals, delimiter) + "\n"
			out.Write([]byte(line))
		}
	}
	return nil
}
