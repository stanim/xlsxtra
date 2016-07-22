package xlsxtra_test

import (
	"fmt"

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
