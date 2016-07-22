package xlsxtra

import (
	"sort"

	"github.com/tealeg/xlsx"
)

// Sort sheet rows according to multi column
func Sort(sheet *xlsx.Sheet, start, end int,
	columns ...int) {
	m := NewMultiColumnSort(sheet, start, end)
	m.Sort(columns...)
}

// MultiColumnSort implements the Sort interface. It
// provides multi-column sort for certain rows of a sheet,
// which are selected by begin and end indices. If End is
// is -1, the last row of the sheet will be selected.
type MultiColumnSort struct {
	Sheet      *xlsx.Sheet
	Columns    []int
	Start, End int
}

// NewMultiColumnSort creates a new multi column sorter.
func NewMultiColumnSort(
	sheet *xlsx.Sheet, start, end int) *MultiColumnSort {
	return &MultiColumnSort{
		Sheet: sheet,
		Start: start,
		End:   end,
	}
}

// Sort executes the multi-column sort of the rows
func (m *MultiColumnSort) Sort(columns ...int) {
	m.Columns = columns
	sort.Sort(m)
}

// Len is part of sort.Interface.
func (m *MultiColumnSort) Len() int {
	end := m.End
	last := len(m.Sheet.Rows) - 1
	if end == -1 || end > last {
		end = last
	}
	return end - m.Start + 1
}

// Swap is part of sort.Interface.
func (m *MultiColumnSort) Swap(i, j int) {
	a := m.Start + i
	b := m.Start + j
	m.Sheet.Rows[a], m.Sheet.Rows[b] =
		m.Sheet.Rows[b], m.Sheet.Rows[a]
}

// get retrieves value by column index, returns empty
// string if doesn't exist.
func get(row *xlsx.Row, col int) string {
	if col < len(row.Cells) {
		s, _ := row.Cells[col].String()
		return s
	}
	return ""
}

// getReverse returns a positive column index and a bool to
// indicate the order should be reserved
func getReverse(index int) (int, bool) {
	reverse := false
	if index < 0 {
		reverse = true
		index = -index
	}
	return index - 1, reverse
}

// Less is part of sort.Interface. It is implemented by
// looping along the indices until it finds a comparison
// that is either Less or !Less.
func (m *MultiColumnSort) Less(i, j int) bool {
	p, q := m.Sheet.Rows[m.Start+i], m.Sheet.Rows[m.Start+j]
	// Try all but the last comparison.
	var c int
	for c = 0; c < len(m.Columns)-1; c++ {
		index, reverse := getReverse(m.Columns[c])
		switch {
		case get(p, index) < get(q, index):
			return !reverse
		case get(q, index) < get(p, index):
			return reverse
		}
		// p == q; try the next comparison.
	}
	// All comparisons to here said "equal", so just return
	// whatever the final comparison reports.
	index, reverse := getReverse(m.Columns[c])
	if get(p, index) < get(q, index) {
		return !reverse
	}
	return reverse
}
