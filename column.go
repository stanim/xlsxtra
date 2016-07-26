package xlsxtra

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

// Col retrieves values by header label from a row
type Col map[string]int

// NewCol creates a new Col from a header row
func NewCol(header *xlsx.Row) Col {
	col := make(Col)
	for i, cell := range header.Cells {
		title, _ := cell.String()
		if title != "" {
			col[title] = i + 1
			col[fmt.Sprintf("-%s", title)] = -col[title]
		}
	}
	return col
}

// Index of a given column header title
func (c Col) Index(title string) (int, error) {
	if i, ok := c[title]; ok {
		return i, nil
	}
	return 0, fmt.Errorf("Unknown column header: %s (%#v)",
		title, c)
}

// Indices of given column header titles
func (c Col) Indices(headers ...string) (
	[]int, error) {
	indices := make([]int, len(headers))
	for i, title := range headers {
		index, err := c.Index(title)
		if err != nil {
			return []int{}, err
		}
		indices[i] = index
	}
	return indices, nil
}

// IndexRow returns index of a given column header title
// inside a row.
func (c Col) IndexRow(
	row *xlsx.Row, title string) (int, error) {
	i, err := c.Index(title)
	if err != nil {
		return 0, err
	}
	if i <= len(row.Cells) {
		return i, nil
	}
	return i, fmt.Errorf(
		`IndexRow: Index %d out of range for column title: %s
Try to set a value in a column to the far right.
[%d][%s]`,
		i, title, len(row.Cells),
		strings.Join(ToString(row), "|"))
}

// Bool value of (row,col) in spreadsheet
func (c Col) Bool(row *xlsx.Row, header string) (bool,
	error) {
	val, err := c.String(row, header)
	if err != nil {
		return false, err
	}
	val = strings.ToLower(val)
	return val == "ja" || val == "yes" || val == "1", nil
}

// BoolMap value of (row,col) in spreadsheet
func (c Col) BoolMap(row *xlsx.Row, headers []string) (
	map[string]bool, error) {
	var err error
	bmap := make(map[string]bool)
	for _, header := range headers {
		bmap[header], err = c.Bool(row, header)
		if err != nil {
			return nil, err
		}
	}
	return bmap, nil
}

// Int value of (row,col) in spreadsheet
func (c Col) Int(row *xlsx.Row, header string) (int,
	error) {
	i, err := c.IndexRow(row, header)
	if err != nil {
		return 0, err
	}
	return row.Cells[i-1].Int()
}

// Float value of (row,col) in spreadsheet
func (c Col) Float(row *xlsx.Row, header string) (float64,
	error) {
	i, err := c.IndexRow(row, header)
	if err != nil {
		return 0, err
	}
	val := row.Cells[i-1].Value
	if strings.HasPrefix(val, "€") ||
		strings.HasPrefix(val, "$") {
		val = strings.TrimLeft(val, "€$ ")
	}
	f, err := strconv.ParseFloat(val, 64)
	if err != nil {
		return math.NaN(), err
	}
	return f, nil
}

// String value of (row,col) in spreadsheet
func (c Col) String(row *xlsx.Row, header string) (string,
	error) {
	i, err := c.IndexRow(row, header)
	if err != nil {
		return "", err
	}
	/** unreachable code fix this for all
		if i >= len(row.Cells) {
			return "", nil
		}
	**/
	return row.Cells[i-1].String()
}

// StringFloatMap converts column with days string into
// a map of floats.
func (c Col) StringFloatMap(row *xlsx.Row, header string,
	dmap map[string]float64, val float64, sep string,
	chars int) error {
	// days
	s, err := c.String(row, header)
	if err != nil {
		return err
	}
	ds := strings.Split(s, sep)
	for _, d := range ds {
		if chars > 0 && len(s) > chars {
			d = d[:chars]
		}
		dmap[d] = val
	}
	return nil
}
