package xlsxtra

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// File extends xlsx.File
type File struct {
	*xlsx.File
	filename string
}

// NewFile creates new spreadsheet file
func NewFile() *File {
	return &File{File: xlsx.NewFile(), filename: "?"}
}

// OpenFile opens a new excel file
func OpenFile(fn string) (*File, error) {
	f, err := xlsx.OpenFile(fn)
	if err != nil {
		return nil, fmt.Errorf("OpenFile: %v", err)
	}
	return &File{File: f, filename: fn}, nil
}

// AddSheet with certain name to spreadsheet file
func (f *File) AddSheet(name string) (*Sheet, error) {
	sheet, err := f.File.AddSheet(name)
	if err != nil {
		return nil, fmt.Errorf("AddSheet: %v", err)
	}
	return &Sheet{Sheet: sheet}, nil
}

// SheetByName get sheet by name from spreadsheet
func (f *File) SheetByName(name string) (*Sheet, error) {
	sheet, ok := f.Sheet[name]
	if !ok {
		return nil, fmt.Errorf(
			"SheetByName(%q): file %q does not contain this sheet",
			name, f.filename)
	}
	return &Sheet{Sheet: sheet}, nil
}

// SheetByIndex get sheet by index from spreadsheet
func (f *File) SheetByIndex(index int) *Sheet {
	return &Sheet{f.Sheets[index]}
}

// SheetRange returns sheet range including end sheet.
// Negative indices can be used.
func (f *File) SheetRange(start, end int) []*Sheet {
	n := len(f.Sheets)
	if start < 0 {
		start += n
	}
	if end <= 0 {
		end += n
	}
	return Sheets(f.Sheets[start : end+1])
}

// SheetMap returns a map of sheets by name
func (f *File) SheetMap() map[string]*Sheet {
	sheetMap := make(map[string]*Sheet)
	for _, sheet := range f.Sheets {
		sheetMap[sheet.Name] = &Sheet{Sheet: sheet}
	}
	return sheetMap
}
