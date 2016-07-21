package xlsxtra_test

import (
	"fmt"
	"testing"

	"github.com/stanim/xlsxtra"
	"github.com/tealeg/xlsx"
)

func ExampleToString() {
	titles := []string{"Rob", "Robert", "Ken"}
	sheet, err := xlsx.NewFile().AddSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	row := sheet.AddRow()
	for _, title := range titles {
		row.AddCell().SetString(title)
	}
	fmt.Printf("%v", xlsxtra.ToString(row))
	// Output:
	// [Rob Robert Ken]
}

func ExampleCol() {
	type Item struct {
		Name   string
		Price  float64
		Amount int
	}
	sheet, err := xlsx.NewFile().AddSheet("Shopping Basket")
	if err != nil {
		fmt.Println(err)
		return
	}
	// column header
	var titles = []string{"item", "price", "amount", "total"}
	header := sheet.AddRow()
	for _, title := range titles {
		xlsxtra.AddString(header, title)
	}
	style := xlsxtra.NewStyle(
		"00ff0000", // color
		&xlsx.Font{Size: 10, Name: "Arial", Bold: true}, // bold
		nil, // border
		nil, // alignment
	)
	xlsxtra.SetRowStyle(header, style)
	// items
	var items = []Item{
		{"chocolate", 4.99, 2},
		{"cookies", 6.45, 3},
	}
	var row *xlsx.Row
	for i, item := range items {
		row = sheet.AddRow()
		xlsxtra.AddString(row, item.Name)
		xlsxtra.AddFloat(row, item.Price, "0.00")
		xlsxtra.AddInt(row, item.Amount)
		xlsxtra.AddFormula(row,
			fmt.Sprintf("B%d*C%d", i+1, i+1), "0.00")
	}
	// column Col type
	col := xlsxtra.NewCol(header)
	price, err := col.Float(row, "price")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(price)
	// Output: 6.45
}

func TestOpenSheet(t *testing.T) {
	sheet, err := xlsxtra.OpenSheet(
		"xlsxtra_test.xlsx", "foo")
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

type Data struct {
	b       bool
	f       float64
	formula string
	i       int
	s       string
}

func TestCol(t *testing.T) {
	var (
		titles = []string{
			"bool", "float", "formula", "int", "string"}
		data = []Data{
			{true, 3.14, "B2*2", 2, "monday, tuesday"},
			{false, 3.14, "B2*2", 2, "€ 2.50"},
		}
	)
	sheet, err := xlsx.NewFile().AddSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	// column header titles
	header := sheet.AddRow()
	for _, title := range titles {
		xlsxtra.AddString(header, title)
	}
	// table data
	for _, d := range data {
		row := sheet.AddRow()
		xlsxtra.AddBool(row, d.b)
		xlsxtra.AddFloat(row, d.f, "0.00")
		xlsxtra.AddFormula(row, d.formula, "0.0")
		xlsxtra.AddInt(row, d.i)
		xlsxtra.AddString(row, d.s)
	}
	row := sheet.AddRow()
	xlsxtra.AddBool(row, false)
	col := xlsxtra.NewCol(header)
	row1 := sheet.Rows[1]
	row2 := sheet.Rows[2]
	row3 := sheet.Rows[3]
	// (bool)map
	bmGot, err := col.BoolMap(row1, []string{"bool"})
	if err != nil {
		t.Fatal(err)
	}
	bGot := bmGot["bool"]
	bWant := data[0].b
	if bGot != bWant {
		t.Fatalf("Col.Bool: got \"%v\"; want \"%v\"",
			bGot, bWant)
	}
	bGot, err = col.Bool(row2, "bool")
	if err != nil {
		t.Fatal(err)
	}
	bWant = false //data[1].b
	if bGot != bWant {
		t.Fatalf("Col.Bool: got \"%v\"; want \"%v\"",
			bGot, bWant)
	}
	// float
	fGot, err := col.Float(row1, "float")
	if err != nil {
		t.Fatal(err)
	}
	fWant := data[0].f
	if fGot != fWant {
		t.Fatalf("Col.Float: got \"%v\"; want \"%v\"",
			fGot, fWant)
	}
	fGot, err = col.Float(row1, "string") // not parseable
	if err == nil {
		t.Fatal("Col.Float: expected invalid syntax for strconv.ParseFloat")
	}
	fGot, err = col.Float(row2, "string") // euro sign
	if err != nil {
		t.Fatal(err)
	}
	fWant = 2.5
	if fGot != fWant {
		t.Fatalf("Col.Float: got \"%v\"; want \"%v\"",
			fGot, fWant)
	}
	// int
	iGot, err := col.Int(row1, "int")
	if err != nil {
		t.Fatal(err)
	}
	iWant := data[0].i
	if iGot != iWant {
		t.Fatalf("Col.Int: got \"%d\"; want \"%d\"",
			iGot, iWant)
	}
	// (string)floatMap
	sfmGot := make(map[string]float64)
	sWant := 1.0
	err = col.StringFloatMap(
		row1, "string", sfmGot, sWant, ", ", 3)
	if err != nil {
		t.Fatal(err)
	}
	sGot := sfmGot["mon"]
	if sGot != sWant {
		t.Fatalf("Col.Float: got \"%v\"; want \"%v\" (%#v)",
			sGot, sWant, sfmGot)
	}
	// out of range
	iGot, err = col.Int(row3, "int")
	if err == nil {
		t.Fatal("col.Index: expected error out of range")
	}
	// not existing
	bmGot, err = col.BoolMap(row1, []string{"not existing"})
	if err == nil {
		t.Fatal("col.BoolMap: expected error for not existing")
	}
	fGot, err = col.Float(row1, "not existing")
	if err == nil {
		t.Fatal("col.Float: expected error for not existing")
	}
	iGot, err = col.Int(row1, "not existing")
	if err == nil {
		t.Fatal("col.Int: expected error for not existing")
	}
	err = col.StringFloatMap(
		row1, "not exisiting", sfmGot, sWant, ", ", 3)
	if err == nil {
		t.Fatal("col.StringFloatMap: expected error for not existing")
	}
	// style
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
	xlsxtra.SetRowStyle(row1, style)
	style = row1.Cells[0].GetStyle()
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
