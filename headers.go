package xlsxtra

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// Headers defines a map of column indices (int) by header
// title (string).
type Headers map[string]int

// NewHeaders creates new Headers from column header
// titles.
func NewHeaders(row *xlsx.Row) Headers {
	hs := make(Headers)
	for i, cell := range row.Cells {
		title, _ := cell.String()
		if title != "" {
			hs[title] = i + 1
			hs[fmt.Sprintf("-%s", title)] = -hs[title]
		}
	}
	return hs
}

// Index of a given column header title
func (hs Headers) Index(title string) (int, error) {
	if i, ok := hs[title]; ok {
		return i, nil
	}
	return 0, fmt.Errorf("Unknown column header: %s (%#v)",
		title, hs)
}

// Indices of given column header titles
func (hs Headers) Indices(titles ...string) (
	[]int, error) {
	indices := make([]int, len(titles))
	for i, title := range titles {
		index, err := hs.Index(title)
		if err != nil {
			return []int{}, err
		}
		indices[i] = index
	}
	return indices, nil
}
