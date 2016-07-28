package xlsxtra_test

import (
	"fmt"
	"testing"

	"github.com/stanim/xlsxtra"
	"github.com/tealeg/xlsx"
)

func TestOpenSheet(t *testing.T) {
	sheet, err := xlsxtra.OpenSheet(
		"xlsxtra_test.xlsx", "sort_test.go")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.Rows[0]
	want := "id"
	got := row.Cells[0].Value
	if got != want {
		t.Errorf("Got %q; want %q in first cell", got, want)
	}
	// Row tests
	rows := sheet.RowRange(-4, -1)
	if len(rows) != 4 {
		t.Fatalf("Got %d rows; expected 4 rows", len(rows))
	}
	if rows[0].Cells[1].Value != "Donald" {
		t.Fatalf("Got %q, expected \"Donald\"",
			rows[0].Cells[0].Value)
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

func newSheetSheet() *xlsxtra.Sheet {
	file := xlsx.NewFile()
	s, _ := file.AddSheet("Sheet")
	sheet := &xlsxtra.Sheet{Sheet: s}
	data := [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}
	for _, r := range data {
		row := sheet.AddRow()
		for _, c := range r {
			row.AddString(c)
		}
	}
	return sheet
}

func TestSheet_Row(t *testing.T) {
	sheet := newSheetSheet()
	row := sheet.Row(1)
	got := row.Cells[0].Value
	want := "A1"
	if got != want {
		t.Fatalf("got %s; want %s", got, want)
	}
}

func ExampleSheet_Cell() {
	file := xlsxtra.NewFile()
	sheet, err := file.AddSheet("Sheet")
	if err != nil {
		fmt.Println(err)
		return
	}
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "I am a cell!"
	cell, err = sheet.Cell("A1")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell.Value)
	// Output: I am a cell!
}

func TestSheet_Cell(t *testing.T) {
	file := xlsxtra.NewFile()
	sheet, err := file.AddSheet("Sheet")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "I am a cell!"
	cell, err = sheet.Cell("A2")
	if err == nil {
		t.Fatal("Expected error: row of \"A2\" out of range")
	}
	cell, err = sheet.Cell("B1")
	if err == nil {
		t.Fatal("Expected error: column of \"B1\" out of range")
	}
	cell, err = sheet.Cell("ZZZZ")
	if err == nil {
		t.Fatal("Expected error: invalid coord \"ZZZZ\"")
	}
	cell, err = sheet.Cell("ZZZZ1")
	if err == nil {
		t.Fatal("Expected error: column \"ZZZZ1\" out of range")
	}
}

func ExampleSheet_CellRange() {
	var print = func(cells [][]*xlsx.Cell) {
		for _, r := range cells {
			fmt.Printf("|")
			for _, c := range r {
				fmt.Printf("%s|", c.Value)
			}
			fmt.Println()
		}
	}
	file := xlsxtra.NewFile()
	sheet, err := file.AddSheet("Sheet")
	if err != nil {
		fmt.Println(err)
		return
	}
	data := [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}
	for _, r := range data {
		row := sheet.AddRow()
		for _, c := range r {
			row.AddString(c)
		}
	}
	cells, err := sheet.CellRange("A1:B2")
	if err != nil {
		fmt.Println(err)
		return
	}
	print(cells)
	fmt.Println()
	print(xlsxtra.Transpose(cells))
	// Output:
	// |A1|B1|
	// |A2|B2|
	//
	// |A1|A2|
	// |B1|B2|
}

func newSheetUtils() *xlsxtra.Sheet {
	file := xlsxtra.NewFile()
	sheet, _ := file.AddSheet("Sheet")
	data := [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}
	for _, r := range data {
		row := sheet.AddRow()
		for _, c := range r {
			row.AddString(c)
		}
	}
	return sheet
}

// TestCellRange corner cases
func TestSheet_CellRange(t *testing.T) {
	sheet := newSheetUtils()
	_, err := sheet.CellRange("A0:B2")
	if err == nil {
		t.Fatal("Expected error as row 0 does not exist")
	}
	_, err = sheet.CellRange("A1:C2")
	if err == nil {
		t.Fatal("Expected error as column C is out of range")
	}
}
