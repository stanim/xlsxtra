package xlsxtra

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
)

const (
	rangeExpr = `[$]?(?P<min_col>[A-Z]+)[$]?(?P<min_row>[1-9]\d*)(:[$]?(?P<max_col>[A-Z]+)[$]?(?P<max_row>[1-9]\d*))?`
)

var (
	// ColStr maps an integer column index to its name
	ColStr = make(map[int]string)
	// StrCol maps a column name to its integer index
	StrCol  = make(map[string]int)
	reCoord = regexp.MustCompile(
		`^[$]?([A-Z]+)[$]?([1-9]\d*)$`)
	reRange = regexp.MustCompile(
		fmt.Sprintf("^%s$", rangeExpr))
)

// SplitCoord splits a coordinate string into column and
// row. (For example "AA19" is split into "AA" & "19")
func SplitCoord(coord string) (string, int, error) {
	m := reCoord.FindStringSubmatch(coord)
	if m == nil {
		return "", 0,
			fmt.Errorf("SplitCoord: Invalid cell coordinates %q",
				coord)
	}
	column, rowStr := m[1], m[2]
	row, _ := strconv.Atoi(rowStr)
	return column, row, nil
}

func colStr(i int) string {
	letters := []string{}
	for i > 0 {
		mod := int(math.Mod(float64(i), 26))
		i = int(float64(i) / 26)
		// check for exact division and borrow if needed
		if mod == 0 {
			mod = 26
			i--
		}
		letters = append(
			[]string{string(mod + 64)}, letters...)
	}
	return strings.Join(letters, "")
}

// RangeBounds converts a range string into boundaries:
// min_col, min_row, max_col, max_row.
// Cell coordinates will be converted into a range with
// the cell at both end.
func RangeBounds(s string) (int, int, int, int, error) {
	m := reRange.FindStringSubmatch(s)
	if m == nil {
		return 0, 0, 0, 0, fmt.Errorf("Invalid range %q", s)
	}
	minCol := StrCol[m[1]]
	minRow, _ := strconv.Atoi(m[2])
	maxCol := StrCol[m[4]]
	maxRow, _ := strconv.Atoi(m[5])
	return minCol, minRow, maxCol, maxRow, nil
}

func init() {
	var col string
	for i := 1; i < 18279; i++ {
		col = colStr(i)
		ColStr[i] = col
		StrCol[col] = i
	}
}
