package xlsxtra

import "github.com/tealeg/xlsx"

// Row of a sheet
type Row struct {
	*xlsx.Row
}

// AddBool adds a cell with bool as 1 or 0 to a row
func (row *Row) AddBool(x ...bool) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		if y {
			cell = row.AddInt(1)
		} else {
			cell = row.AddInt(0)
		}
	}
	return cell
}

// AddEmpty adds n empty cells to a row
func (row *Row) AddEmpty(n int) {
	for i := 0; i < n; i++ {
		row.AddCell()
	}
}

// AddFloat adds a cell with float64 value to a row
func (row *Row) AddFloat(format string, x ...float64,
) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell = row.AddCell()
		cell.SetFloatWithFormat(y, format)
	}
	return cell
}

// AddFormula adds a cell with formula to a row
func (row *Row) AddFormula(format string,
	formula ...string) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range formula {
		cell = row.AddCell()
		cell.SetFormula(y)
		cell.NumFmt = format
	}
	return cell
}

// AddInt adds a cell with int value to a row
func (row *Row) AddInt(x ...int) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell = row.AddCell()
		cell.SetInt(y)
	}
	return cell
}

// AddString adds a cell with string value to a row
func (row *Row) AddString(x ...string) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell = row.AddCell()
		cell.SetString(y)
	}
	return cell
}

// SetStyle set style to all cells of a row
func (row *Row) SetStyle(style *xlsx.Style) {
	for _, cell := range row.Cells {
		cell.SetStyle(style)
	}
}

// ToString converts row to string slice
func ToString(cells []*xlsx.Cell) []string {
	s := make([]string, len(cells))
	for i, cell := range cells {
		s[i] = cell.Value
	}
	return s
}

// Rows converts slice of xlsx.Row into Row
func Rows(rows []*xlsx.Row) []*Row {
	r := make([]*Row, len(rows))
	for i, row := range rows {
		r[i] = &Row{row}
	}
	return r
}
