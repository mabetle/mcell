package wxlsx

import (
	"github.com/mabetle/mcore/mtest"
	"testing"
)

func TestGetCellValue(t *testing.T) {
	mtest.RegTest(t)
	fn := "testdata/demo.xlsx"
	in := GetCellValue(fn, "Sheet1", "A1", "")
	//println(in)
	mtest.AssertEqual(in, "ID")
}
