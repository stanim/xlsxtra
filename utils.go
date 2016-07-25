package xlsxtra

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	maxCol    = 18279
	rangeExpr = `[$]?([A-Z]+)[$]?([1-9]\d*)(:[$]?([A-Z]+)[$]?([1-9]\d*))?`
)

var (
	// ColStr maps an integer column index to its name
	ColStr [maxCol]string
	// StrCol maps a column name to its integer index
	StrCol  = make(map[string]int)
	reCoord = regexp.MustCompile(
		`^[$]?([A-Z]+)[$]?([1-9]\d*)$`)
	reRange = regexp.MustCompile(
		fmt.Sprintf("^%s$", rangeExpr))
)

// ColRange gives a range of intervals.
// (Returns empty slice for invalid input.)
func ColRange(start, end string) []string {
	s := StrCol[start]
	e := StrCol[end] + 1
	if s == 0 || e == 1 {
		return nil
	}
	r := make([]string, e-s)
	for i := s; i < e; i++ {
		r[i-s] = ColStr[i]
	}
	return r
}

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

func init() {
	var col string
	ColStr[0] = "?"
	for i := 1; i < maxCol; i++ {
		col = colStr(i)
		ColStr[i] = col
		StrCol[col] = i
	}
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

// Abs converts a coordinate to an absolute coordinate
// (An invalid string is returned unaltered.)
func Abs(s string) string {
	m := reRange.FindStringSubmatch(s)
	if m == nil {
		return s
	}
	if m[4] == "" || m[5] == "" {
		return fmt.Sprintf("$%s$%s", m[1], m[2])
	}
	return fmt.Sprintf("$%s$%s:$%s$%s",
		m[1], m[2], m[4], m[5])
}

// RangeBounds converts a range string into boundaries:
// min_col, min_row, max_col, max_row.
// Cell coordinates will be converted into a range with
// the cell at both end.
func RangeBounds(rg string) (int, int, int, int, error) {
	m := reRange.FindStringSubmatch(rg)
	if m == nil {
		return 0, 0, 0, 0, fmt.Errorf("Invalid range %q", rg)
	}
	minCol := StrCol[m[1]]
	minRow, _ := strconv.Atoi(m[2])
	var maxCol, maxRow int // m[4], m[5]
	if m[4] != "" && m[5] != "" {
		maxCol = StrCol[m[4]]
		maxRow, _ = strconv.Atoi(m[5])
	} else {
		maxCol = minCol
		maxRow = minRow
	}
	return minCol, minRow, maxCol, maxRow, nil
}

func checkCell(sheet *xlsx.Sheet, col, row int) (
	*xlsx.Row, error) {
	n := len(sheet.Rows)
	if row > n {
		return nil, fmt.Errorf(
			"checkCell: row %d out of range of sheet (max %d)",
			row, n)
	}
	r := sheet.Rows[row-1]
	n = len(r.Cells)
	if col > n {
		return nil, fmt.Errorf(
			"checkCell: column %d out of range (max %d)",
			col, n)
	}
	return r, nil
}

// GetCell returns a cell based on coordinate string.
func GetCell(sheet *xlsx.Sheet, coord string) (
	*xlsx.Cell, error) {
	colS, row, err := SplitCoord(coord)
	if err != nil {
		return nil, fmt.Errorf("checkCell: %v", err)
	}
	col, ok := StrCol[colS]
	if !ok {
		return nil, fmt.Errorf("checkCell: column %q overflow",
			colS)
	}
	r, err := checkCell(sheet, col, row)
	if err != nil {
		return nil, fmt.Errorf("GetCell: %v", err)
	}
	return r.Cells[col-1], nil
}

// GetCellRange returns all cells by row
func GetCellRange(sheet *xlsx.Sheet, rg string) (
	[][]*xlsx.Cell, error) {
	minCol, minRow, maxCol, maxRow, err := RangeBounds(rg)
	if err != nil {
		return nil, fmt.Errorf("GetCellRange: %v", err)
	}
	_, err = checkCell(sheet, maxCol, maxRow)
	if err != nil {
		return nil, fmt.Errorf("GetCellRange: %v", err)
	}
	rows := sheet.Rows
	nRow := maxRow - minRow + 1
	nCol := maxCol - minCol + 1
	result := make([][]*xlsx.Cell, nRow)
	for r := minRow; r <= maxRow; r++ {
		row := rows[r-1]
		cells := make([]*xlsx.Cell, nCol)
		for col := minCol; col <= maxCol; col++ {
			cells[col-minCol] = row.Cells[col-1]
		}
		result[r-minRow] = cells
	}
	return result, nil
}

// Transpose rows into columns and vice versa
func Transpose(cells [][]*xlsx.Cell) [][]*xlsx.Cell {
	rows := len(cells)
	if rows == 0 {
		return cells
	}
	cols := len(cells[0])
	result := make([][]*xlsx.Cell, cols)
	for c := 0; c < cols; c++ {
		cs := make([]*xlsx.Cell, rows)
		for r := 0; r < rows; r++ {
			cs[r] = cells[r][c]
		}
		result[c] = cs
	}
	return result
}
