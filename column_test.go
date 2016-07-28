package xlsxtra_test

import (
	"fmt"
	"testing"

	"github.com/stanim/xlsxtra"
	"github.com/tealeg/xlsx"
)

func ExampleCol() {
	type Item struct {
		Name   string
		Price  float64
		Amount int
	}
	sheet, err := xlsxtra.NewFile().AddSheet("Basket")
	if err != nil {
		fmt.Println(err)
		return
	}
	// column header
	var headers = []string{"item", "price", "amount", "total"}
	header := sheet.AddRow()
	for _, title := range headers {
		header.AddString(title)
	}
	style := xlsxtra.NewStyle(
		"00ff0000", // color
		&xlsx.Font{Size: 10, Name: "Arial", Bold: true}, // bold
		nil, // border
		nil, // alignment
	)
	header.SetStyle(style)
	// items
	var items = []Item{
		{"chocolate", 4.99, 2},
		{"cookies", 6.45, 3},
	}
	var row *xlsxtra.Row
	for i, item := range items {
		row = sheet.AddRow()
		row.AddString(item.Name)
		row.AddFloat("0.00", item.Price)
		row.AddInt(item.Amount)
		row.AddFormula("0.00",
			fmt.Sprintf("B%d*C%d", i+1, i+1))
	}
	// column Col type
	col := xlsxtra.NewCol(sheet, 1)
	price, err := col.Float(row, "price")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(price)
	// Output: 6.45
}

type Data struct {
	b       bool
	f       float64
	formula string
	i       int
	s       string
}

func newSheet(t *testing.T, headers []string,
	data []Data) *xlsxtra.Sheet {
	sheet, err := xlsxtra.NewFile().AddSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	// column header titles
	header := sheet.AddRow()
	for _, title := range headers {
		header.AddString(title)
	}
	// table data
	for _, d := range data {
		row := sheet.AddRow()
		row.AddBool(d.b)
		row.AddFloat("0.00", d.f)
		row.AddFormula("0.00", d.formula)
		row.AddInt(d.i)
		row.AddString(d.s)
		row.AddEmpty(2)
	}
	row := sheet.AddRow()
	row.AddBool(false)
	return sheet
}

func checkBool(t *testing.T, data []Data, col xlsxtra.Col,
	row1, row2 *xlsxtra.Row) {
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
}

func checkFloat(t *testing.T, data []Data, col xlsxtra.Col,
	row1, row2 *xlsxtra.Row) {
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
}

func checkInt(t *testing.T, data []Data, col xlsxtra.Col,
	row1, row2 *xlsxtra.Row) {
	iGot, err := col.Int(row1, "int")
	if err != nil {
		t.Fatal(err)
	}
	iWant := data[0].i
	if iGot != iWant {
		t.Fatalf("Col.Int: got \"%d\"; want \"%d\"",
			iGot, iWant)
	}
	iGot, err = col.Int(row2, "int")
	if err != nil {
		t.Fatal(err)
	}
	iWant = data[1].i
	if iGot != iWant {
		t.Fatalf("Col.Int: got \"%d\"; want \"%d\"",
			iGot, iWant)
	}
}

func checkString(t *testing.T, data []Data,
	col xlsxtra.Col, row1 *xlsxtra.Row) {
	// (string)floatMap
	sfmGot := make(map[string]float64)
	sWant := 1.0
	err := col.StringFloatMap(
		row1, "string", sfmGot, sWant, ", ", 3)
	if err != nil {
		t.Fatal(err)
	}
	sGot := sfmGot["mon"]
	if sGot != sWant {
		t.Fatalf(
			"Col.StringFloatMap: got \"%v\"; want \"%v\" (%#v)",
			sGot, sWant, sfmGot)
	}
}

func checkEmpty(t *testing.T, col xlsxtra.Col,
	row1 *xlsxtra.Row) {
	for _, empty := range []string{"empty1", "empty2"} {
		sGot, err := col.String(row1, empty)
		if err != nil {
			t.Fatal(err)
		}
		sWant := ""
		if sGot != sWant {
			t.Fatalf("Empty: got %q; want %q",
				sGot, sWant)
		}
	}
}

func checkErrors(t *testing.T, data []Data,
	col xlsxtra.Col, row1, row3 *xlsxtra.Row) {
	sfmGot := make(map[string]float64)
	sWant := 1.0
	// out of range
	_, err := col.Int(row3, "int")
	if err == nil {
		t.Fatal("col.Index: expected error out of range")
	}
	// not existing
	_, err = col.BoolMap(row1, []string{"not existing"})
	if err == nil {
		t.Fatal("col.BoolMap: expected error for not existing")
	}
	_, err = col.Float(row1, "not existing")
	if err == nil {
		t.Fatal("col.Float: expected error for not existing")
	}
	_, err = col.Int(row1, "not existing")
	if err == nil {
		t.Fatal("col.Int: expected error for not existing")
	}
	err = col.StringFloatMap(
		row1, "not exisiting", sfmGot, sWant, ", ", 3)
	if err == nil {
		t.Fatal("col.StringFloatMap: expected error for not existing")
	}
}

func TestCol(t *testing.T) {
	var (
		headers = []string{
			"bool", "float", "formula", "int", "string",
			"empty1", "empty2"}
		data = []Data{
			{true, 3.14, "B2*2", 2, "monday, tuesday"},
			{false, 3.14, "B2*2", 2, "€ 2.50"},
		}
	)
	// create sheet
	sheet := newSheet(t, headers, data)
	row1 := sheet.Row(2)
	row2 := sheet.Row(3)
	row3 := sheet.Row(4)
	col := xlsxtra.NewCol(sheet, 1)
	// check
	checkBool(t, data, col, row1, row2)
	checkFloat(t, data, col, row1, row2)
	checkInt(t, data, col, row1, row2)
	checkString(t, data, col, row1)
	checkEmpty(t, col, row1)
	checkErrors(t, data, col, row1, row3)
}
