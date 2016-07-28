package xlsxtra

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// Sheet extends xlsx.Sheet
type Sheet struct {
	*xlsx.Sheet
}

// OpenSheet open a sheet from an xlsx file. If you need
// to use multiple sheets from one file use the Sheets type
// instead.
func OpenSheet(fn string, name string) (
	*Sheet, error) {
	f, err := OpenFile(fn)
	if err != nil {
		return nil, fmt.Errorf("OpenSheet: %v", err)
	}
	return f.SheetByName(name)
}

// AddRow adds a row to a sheet
func (sheet *Sheet) AddRow() *Row {
	return &Row{Row: sheet.Sheet.AddRow()}
}

// Row returns one based row
func (sheet *Sheet) Row(row int) *Row {
	return &Row{Row: sheet.Rows[row-1]}
}

// RowRange return all rows.
func (sheet *Sheet) RowRange(start, end int) []*Row {
	n := len(sheet.Rows)
	if start < 0 {
		start += n + 1
	}
	if end <= 0 {
		end += n + 1
	}
	start--
	return Rows(sheet.Rows[start:end])
}

func (sheet *Sheet) checkCell(col, row int) (
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

// Cell returns a cell based on coordinate string.
func (sheet *Sheet) Cell(coord string) (
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
	r, err := sheet.checkCell(col, row)
	if err != nil {
		return nil, fmt.Errorf("Cell: %v", err)
	}
	return r.Cells[col-1], nil
}

// CellRange returns all cells by row
func (sheet *Sheet) CellRange(rg string) (
	[][]*xlsx.Cell, error) {
	minCol, minRow, maxCol, maxRow, err := RangeBounds(rg)
	if err != nil {
		return nil, fmt.Errorf("CellRange: %v", err)
	}
	_, err = sheet.checkCell(maxCol, maxRow)
	if err != nil {
		return nil, fmt.Errorf("CellRange: %v", err)
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

// Sheets converts slice of xlsx.Sheet into Sheet
func Sheets(sheets []*xlsx.Sheet) []*Sheet {
	s := make([]*Sheet, len(sheets))
	for i, sheet := range sheets {
		s[i] = &Sheet{sheet}
	}
	return s
}
