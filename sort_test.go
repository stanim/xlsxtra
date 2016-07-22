package xlsxtra_test

import (
	"fmt"
	"strings"

	"github.com/stanim/xlsxtra"
)

// ExampleOpenSheet demonstrates multi column sort
func ExampleSort() {
	sheet, err := xlsxtra.OpenSheet(
		"xlsxtra_test.xlsx", "sort_test.go")
	if err != nil {
		fmt.Println(err)
		return
	}
	// add incomplete row for testing purposes
	row := sheet.AddRow()
	xlsxtra.AddString(row, "11", "Fred", "Bryant")
	// multi column sort
	xlsxtra.Sort(sheet, 1, -1,
		3,  // last name
		-2, // first name
		7,
		6, // ip address
	)
	for _, row := range sheet.Rows {
		fmt.Println(strings.Join(xlsxtra.ToString(row), ", "))
	}
	// Output:
	// id, first_name, last_name, email, gender, ip_address
	// 11, Fred, Bryant
	// 9, Donald, Bryant, lharper8@wunderground.com, Female, 171.36.193.204
	// 7, Donald, Bryant, dbryant6@redcross.org, Male, 57.28.4.238
	// 10, Donald, Bryant, hmarshall9@stumbleupon.com, Male, 74.236.54.34
	// 4, Teresa, Hunter, thall3@arizona.edu, Female, 94.78.214.245
	// 5, Joshua, Hunter, jstone4@google.cn, Male, 125.32.19.210
	// 8, Jacqueline, Hunter, jfields7@dagondesign.com, Female, 165.182.190.97
	// 2, Harry, Hunter, hhunter1@webnode.com, Male, 240.42.189.119
	// 6, Rose, Spencer, rjohnson5@odnoklassniki.ru, Female, 29.139.205.214
	// 1, Jimmy, Spencer, jspencer0@cnet.com, Male, 16.17.167.238
	// 3, Benjamin, Spencer, bmorgan2@unblog.fr, Male, 96.41.142.121
}
