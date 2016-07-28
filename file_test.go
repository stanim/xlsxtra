package xlsxtra_test

import (
	"testing"

	"github.com/stanim/xlsxtra"
)

var sheetNames = []string{"sheet_test.go", "sort_test.go"}

func newFile(t *testing.T) *xlsxtra.File {
	f := xlsxtra.NewFile()
	for _, name := range sheetNames {
		_, err := f.AddSheet(name)
		if err != nil {
			t.Fatal(err)
		}
	}
	return f
}

func TestFile_AddSheet(t *testing.T) {
	f := xlsxtra.NewFile()
	_, _ = f.AddSheet("name")
	_, err := f.AddSheet("name")
	if err == nil {
		t.Fatalf(
			"Adding 2 sheets with same name should give an error")
	}
}

func TestFile_SheetByIndex(t *testing.T) {
	f := xlsxtra.NewFile()
	_, err := f.AddSheet("name")
	if err != nil {
		t.Fatal(err)
	}
	sheet := f.SheetByIndex(0)
	if sheet.Name != "name" {
		t.Fatalf("got %q; want \"name\"", sheet.Name)
	}
}

func TestFile_SheetRange(t *testing.T) {
	f := newFile(t)
	sheets := f.SheetRange(-2, -2)
	if len(sheets) != 1 {
		t.Fatal("Expected only one sheet")
	}
	if sheets[0].Name != sheetNames[0] {
		t.Fatalf("got %q; want %q",
			sheets[0].Name, sheetNames[0])
	}
}

func TestFile_SheetMap(t *testing.T) {
	f := newFile(t)
	sheetMap := f.SheetMap()
	for _, name := range sheetNames {
		_, ok := sheetMap[name]
		if !ok {
			t.Fatalf("Expected sheet %q", name)
		}
	}
}
