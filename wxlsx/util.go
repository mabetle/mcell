package wxlsx

import (
	"math"
	"strconv"
	"strings"
)

// GetLetterIndex returns Excel Cell ColumnIndex
// index start from 0
func GetLetterIndex(letter string) (r int) {
	letter = strings.ToUpper(letter)
	b := []byte(letter)
	for i := 0; i < len(b); i++ {
		seq := int(b[i]) - 64

		if i == len(b)-1 {
			r = seq + r
		} else {
			bNum := seq * int(math.Pow(26, float64(len(b)-i-1)))
			r = bNum + r
		}
	}
	r = r - 1
	return
}

// GetRowColIndex returns Cell row and column index.
// the index start from 0
func GetRowColIndex(cell string) (row, col int) {
	num := "1234567890"
	chars := []byte(cell)
	for index, v := range chars {
		if strings.ContainsAny(num, string(v)) {
			colP := string(chars[:index])
			rowP := string(chars[index:])
			col = GetLetterIndex(colP)
			rowN, err := strconv.ParseInt(rowP, 10, 32)
			if err != nil {
				return
			}
			row = int(rowN) - 1
			return
		}
	}
	return
}
