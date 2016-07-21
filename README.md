# xlsxtra - extra utilities for [xlsx](https://github.com/tealeg/xlsx)

[![Travis CI](https://img.shields.io/travis/stanim/xlsxtra/master.svg?style=flat-square)](https://travis-ci.org/stanim/xlsxtra)
[![codecov](https://codecov.io/gh/stanim/xlsxtra/branch/master/graph/badge.svg)](https://codecov.io/gh/stanim/xlsxtra)
[![goreportcard](https://goreportcard.com/badge/github.com/stanim/xlsxtra)](https://goreportcard.com/report/github.com/stanim/xlsxtra)
[![Documentation and Examples](https://godoc.org/github.com/stanim/xlsxtra?status.svg)](https://godoc.org/github.com/stanim/xlsxtra)
[![Software License](https://img.shields.io/badge/license-mit-orange.svg?style=flat-square)](https://github.com/stanim/xlsxtra/blob/master/LICENSE)

This was developed as an extension for the
[xlsx](https://github.com/tealeg/xlsx)
package. It contains the following utilities to manipulate 
excel files:

- `AddBool()`, `AddInt()`, `AddFloat()`, ...: shortcut to add a cell to a row with the right type.
- `NewStyle()`: create a style and set the `ApplyFill`, `ApplyFont`, `ApplyBorder` and `ApplyAlignment` automatically.
- `NewStyles()`: create a slice of styles based on a color palette
- `Sheets`: access sheets by name instead of by index
- `Col`: access cell values of a row by column header title
- `SetRowStyle`: set style of all cells in a row
- `ToString`: convert a xlsx.Row to a slice of strings

### Example

```go
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
```


### Documentation

See [godoc](https://godoc.org/github.com/stanim/xlsxtra) for more documentation and examples.

### License

Released under the [MIT License](https://github.com/stanim/xlsxtra/blob/master/LICENSE).
