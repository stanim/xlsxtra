package xlsxtra

import "github.com/tealeg/xlsx"

// ToString converts row to string slice
func ToString(row *xlsx.Row) []string {
	s := make([]string, len(row.Cells))
	for i, cell := range row.Cells {
		s[i] = cell.Value
	}
	return s
}

// AddBool adds a cell with bool as 1 or 0 to a row
func AddBool(row *xlsx.Row, x bool) *xlsx.Cell {
	var cell *xlsx.Cell
	if x {
		cell = AddInt(row, 1)
	} else {
		cell = AddInt(row, 0)
	}
	return cell
}

// AddFloat adds a cell with float64 value to a row
func AddFloat(row *xlsx.Row, x float64,
	format string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetFloatWithFormat(x, format)
	return cell
}

// AddFormula adds a cell with formula to a row
func AddFormula(row *xlsx.Row, formula string,
	format string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetFormula(formula)
	cell.NumFmt = format
	return cell
}

// AddInt adds a cell with int value to a row
func AddInt(row *xlsx.Row, x int) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetInt(x)
	return cell
}

// AddString adds a cell with string value to a row
func AddString(row *xlsx.Row, x string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetString(x)
	return cell
}
