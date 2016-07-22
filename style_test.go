package xlsxtra_test

import (
	"testing"

	"github.com/stanim/xlsxtra"
	"github.com/tealeg/xlsx"
)

func checkStyle(t *testing.T, row *xlsx.Row) {
	style := xlsxtra.NewStyle(
		"", // color
		&xlsx.Font{Size: 10, Name: "Arial", Bold: true},
		xlsx.NewBorder("thin", "thin", "thin", "thin"),
		&xlsx.Alignment{
			Horizontal:   "general",
			Indent:       0,
			ShrinkToFit:  false,
			TextRotation: 0,
			Vertical:     "top",
			WrapText:     false,
		},
	)
	if style.ApplyFill {
		t.Fatal("NewStyle: ApplyFill not expected")
	}
	if !style.ApplyFont {
		t.Fatal("NewStyle: ApplyFont expected")
	}
	if !style.ApplyBorder {
		t.Fatal("NewStyle: ApplyBorder expected")
	}
	if !style.ApplyAlignment {
		t.Fatal("NewStyle: ApplyAlignment expected")
	}
	xlsxtra.SetRowStyle(row, style)
	for _, cell := range row.Cells {
		style = cell.GetStyle()
		if style.ApplyFill {
			t.Fatal("NewStyle: ApplyFill not expected")
		}
		if !style.ApplyFont {
			t.Fatal("NewStyle: ApplyFont expected")
		}
		if !style.ApplyBorder {
			t.Fatal("NewStyle: ApplyBorder expected")
		}
		if !style.ApplyAlignment {
			t.Fatal("NewStyle: ApplyAlignment expected")
		}
	}
}

func TestNewStyle(t *testing.T) {
	sheet, err := xlsx.NewFile().AddSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.AddRow()
	xlsxtra.AddString(row, "foo")
	xlsxtra.AddInt(row, 2)
	checkStyle(t, row)
}

func TestNewStyles(t *testing.T) {
	rgb := []string{"00ff0000", "0000ff00", "000000ff"}
	styles := xlsxtra.NewStyles(rgb, nil, nil, nil)
	sGot := styles[0].Fill.FgColor
	sWant := rgb[0]
	if sGot != sWant {
		t.Fatalf("NewStyles: got %q; want %q", sGot, sWant)
	}
	bGot := styles[0].ApplyFill
	bWant := true
	if bGot != bWant {
		t.Fatalf(
			"NewStyles: ApplyFill got \"%v\"; want \"%v\"",
			bGot, bWant)
	}
}
