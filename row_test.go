package xlsxtra_test

import (
	"fmt"

	"github.com/stanim/xlsxtra"
	"github.com/tealeg/xlsx"
)

func ExampleToString() {
	headers := []string{"Rob", "Robert", "Ken"}
	sheet, err := xlsx.NewFile().AddSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	row := sheet.AddRow()
	for _, title := range headers {
		row.AddCell().SetString(title)
	}
	fmt.Printf("%v", xlsxtra.ToString(row.Cells))
	// Output:
	// [Rob Robert Ken]
}
