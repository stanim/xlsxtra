package xlsxtra

import "github.com/tealeg/xlsx"

// AddBool adds a cell with bool as 1 or 0 to a row
func AddBool(row *xlsx.Row, x ...bool) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		if y {
			cell = AddInt(row, 1)
		} else {
			cell = AddInt(row, 0)
		}
	}
	return cell
}

// AddFloat adds a cell with float64 value to a row
func AddFloat(row *xlsx.Row, format string, x ...float64,
) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell := row.AddCell()
		cell.SetFloatWithFormat(y, format)
	}
	return cell
}

// AddFormula adds a cell with formula to a row
func AddFormula(row *xlsx.Row, format string,
	formula ...string) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range formula {
		cell := row.AddCell()
		cell.SetFormula(y)
		cell.NumFmt = format
	}
	return cell
}

// AddInt adds a cell with int value to a row
func AddInt(row *xlsx.Row, x ...int) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell = row.AddCell()
		cell.SetInt(y)
	}
	return cell
}

// AddString adds a cell with string value to a row
func AddString(row *xlsx.Row, x ...string) *xlsx.Cell {
	var cell *xlsx.Cell
	for _, y := range x {
		cell = row.AddCell()
		cell.SetString(y)
	}
	return cell
}

// ToString converts row to string slice
func ToString(row *xlsx.Row) []string {
	s := make([]string, len(row.Cells))
	for i, cell := range row.Cells {
		s[i] = cell.Value
	}
	return s
}
