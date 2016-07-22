package xlsxtra

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// Sheets retrieves a sheet by name
type Sheets struct {
	sheets   map[string]*xlsx.Sheet
	filename string
}

// NewSheets creates a map of sheets by name of an already
// opened xlsx.File
func NewSheets(file *xlsx.File, fn string) *Sheets {
	tabs := Sheets{
		sheets:   make(map[string]*xlsx.Sheet),
		filename: fn,
	}
	for _, sheet := range file.Sheets {
		tabs.sheets[sheet.Name] = sheet
	}
	return &tabs
}

// OpenSheets opens an xlsx file from disk and returns its
// sheets by name
func OpenSheets(fn string) (*Sheets, error) {
	xlFile, err := xlsx.OpenFile(fn)
	if err != nil {
		return nil, err
	}
	return NewSheets(xlFile, fn), nil
}

// Get returns a certain sheet by name. It returns an
// error if a sheet does not exist.
func (t Sheets) Get(name string) (*xlsx.Sheet, error) {
	sheet, ok := t.sheets[name]
	if !ok {
		return nil, fmt.Errorf(
			"file %q does not contain sheet %q",
			t.filename, name)
	}
	return sheet, nil
}

// OpenSheet open a sheet from an xlsx file. If you need
// to use multiple sheets from one file use the Sheets type
// instead.
func OpenSheet(fn string, name string) (
	*xlsx.Sheet, error) {
	tabs, err := OpenSheets(fn)
	if err != nil {
		return nil, err
	}
	return tabs.Get(name)
}
