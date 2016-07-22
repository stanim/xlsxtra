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

- `Sort()`, `SortByHeaders`: multi-column (reverse) sort of selected rows (Note that columns are one based, not zero based to make reverse sort possible.)
- `AddBool()`, `AddInt()`, `AddFloat()`, ...: shortcut to add a cell to a row with the right type.
- `NewStyle()`: create a style and set the `ApplyFill`, `ApplyFont`, `ApplyBorder` and `ApplyAlignment` automatically.
- `NewStyles()`: create a slice of styles based on a color palette
- `Sheets`: access sheets by name instead of by index
- `Col`: access cell values of a row by column header title
- `SetRowStyle`: set style of all cells in a row
- `ToString`: convert a xlsx.Row to a slice of strings

### Example

Add cells and retrieve cell values by column title header:
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

Multi column sort:
```go
sheet, err := xlsxtra.OpenSheet(
	"xlsxtra_test.xlsx", "sort_test.go")
if err != nil {
	fmt.Println(err)
	return
}

// multi column sort
xlsxtra.Sort(sheet, 1, -1,
	3,  // last name
	-2, // first name
	6, // ip address
)
for _, row := range sheet.Rows {
	fmt.Println(strings.Join(xlsxtra.ToString(row), ", "))
}
fmt.Println()
// by header
headers := xlsxtra.NewHeaders(sheet.Rows[0])
err = xlsxtra.SortByHeaders(sheet, 1, -1, headers,
	"-amount",
	"first_name",
)
if err != nil {
	fmt.Println(err)
	return
}
for _, row := range sheet.Rows {
	fmt.Println(strings.Join(xlsxtra.ToString(row), ", "))
}

// Output:
	// Output:
	// id, first_name, last_name, email, gender, amount
	// 9, Donald, Bryant, lharper8@wunderground.com, Female, 100000000
	// 7, Donald, Bryant, dbryant6@redcross.org, Male, 3000000
	// 10, Donald, Bryant, hmarshall9@stumbleupon.com, Male, € 9
	// 4, Teresa, Hunter, thall3@arizona.edu, Female, 6000
	// 5, Joshua, Hunter, jstone4@google.cn, Male, 50000
	// 8, Jacqueline, Hunter, jfields7@dagondesign.com, Female, 20000000
	// 2, Harry, Hunter, hhunter1@webnode.com, Male, 80
	// 6, Rose, Spencer, rjohnson5@odnoklassniki.ru, Female, 400000
	// 1, Jimmy, Spencer, jspencer0@cnet.com, Male, 9
	// 3, Benjamin, Spencer, bmorgan2@unblog.fr, Male, 700
	//
	// id, first_name, last_name, email, gender, amount
	// 9, Donald, Bryant, lharper8@wunderground.com, Female, 100000000
	// 8, Jacqueline, Hunter, jfields7@dagondesign.com, Female, 20000000
	// 7, Donald, Bryant, dbryant6@redcross.org, Male, 3000000
	// 6, Rose, Spencer, rjohnson5@odnoklassniki.ru, Female, 400000
	// 5, Joshua, Hunter, jstone4@google.cn, Male, 50000
	// 4, Teresa, Hunter, thall3@arizona.edu, Female, 6000
	// 3, Benjamin, Spencer, bmorgan2@unblog.fr, Male, 700
	// 2, Harry, Hunter, hhunter1@webnode.com, Male, 80
	// 10, Donald, Bryant, hmarshall9@stumbleupon.com, Male, € 9
	// 1, Jimmy, Spencer, jspencer0@cnet.com, Male, 9
```

### Documentation

See [godoc](https://godoc.org/github.com/stanim/xlsxtra) for more documentation and examples.

### License

Released under the [MIT License](https://github.com/stanim/xlsxtra/blob/master/LICENSE).
