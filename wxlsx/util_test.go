package wxlsx

import (
	"github.com/mabetle/mcore/mtest"
	"testing"
)

func TestGetLetterIndex(t *testing.T) {
	mtest.RegTest(t)
	mtest.AssertEqual(GetLetterIndex("a"), 0)
	mtest.AssertEqual(GetLetterIndex("aa"), 26)
	mtest.AssertEqual(GetLetterIndex("zz"), 25+26*26)
}

func TestGetRowColIndex(t *testing.T) {
	mtest.RegTest(t)
	row, col := GetRowColIndex("AA23")
	mtest.AssertEqual(col, 26, "AA23 row should be 26")
	mtest.AssertEqual(row, 22, "AA23 col should be 22")
}
