package xlsxtra_test

import (
	"testing"

	"github.com/stanim/xlsxtra"
)

func TestOpenSheet(t *testing.T) {
	sheet, err := xlsxtra.OpenSheet(
		"xlsxtra_test.xlsx", "sheet_test.go")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.Rows[0]
	want := "hello"
	got := row.Cells[0].Value
	if got != want {
		t.Errorf("Got %q; want %q in first cell", got, want)
	}
	// file does not exist
	sheet, err = xlsxtra.OpenSheet(
		"xlsxtra_not_existing.xlsx", "foo")
	if err == nil {
		t.Error("Error missing for opening not existing file")
	}
	// sheet does not exist
	sheet, err = xlsxtra.OpenSheet(
		"xlsxtra_test.xlsx", "not exist")
	if err == nil {
		t.Error("Error missing for opening not existing sheet")
	}
}
