package main

import (
	"fmt"
	"math"
)

func GetByteInt(b byte) int {
	return int(b) - 64
}

func GetLetterIndex(letter string) (r int) {
	b := []byte(letter)
	for i := 0; i < len(b); i++ {
		if i == len(b)-1 {
			r = GetByteInt(b[i]) + r
		} else {
			bNum := GetByteInt(b[i]) * int(math.Pow(26, float64(len(b)-i-1)))
			r = bNum + r
		}
	}
	return
}

func main() {
	fmt.Printf("GetLetterIndex(BQ) = %d, should be %d\n", GetLetterIndex("BQ"), 69)
	fmt.Printf("GetLetterIndex(AEX) = %d, should be %d\n", GetLetterIndex("AEX"), 830)
}
