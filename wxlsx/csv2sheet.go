package wxlsx

import (
	"encoding/csv"
	"github.com/tealeg/xlsx"
	"os"
)

// ImportCSV func
func ImportCSV(csvPath string, delimiter string) (*xlsx.File, error) {
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return nil, err
	}
	defer csvFile.Close()
	reader := csv.NewReader(csvFile)
	if len(delimiter) > 0 {
		reader.Comma = rune(delimiter[0])
	} else {
		reader.Comma = rune(',')
	}
	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet(csvPath)
	if err != nil {
		return nil, err
	}
	fields, err := reader.Read()
	for err == nil {
		row := sheet.AddRow()
		for _, field := range fields {
			cell := row.AddCell()
			cell.Value = field
		}
		fields, err = reader.Read()
	}
	// parse csv file error
	if err != nil {
		return nil, err
	}
	// do save
	return xlsxFile, nil
}
