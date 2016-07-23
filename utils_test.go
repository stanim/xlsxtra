package xlsxtra_test

import (
	"fmt"
	"testing"

	"github.com/stanim/xlsxtra"
)

func TestSplitCoord(t *testing.T) {
	column, row, err := xlsxtra.SplitCoord("AA11")
	if err != nil {
		t.Fatal(err)
	}
	if column != "AA" || row != 11 {
		t.Fatalf("expected \"AA\" and 11; got %q and %d",
			column, row)
	}
	_, _, err = xlsxtra.SplitCoord("A0")
	if err == nil {
		t.Fatal("expected error")
	}
}

func Example() {
	fmt.Println(xlsxtra.ColStr[26], xlsxtra.StrCol["AA"])
	// Output: Z 27
}

func ExampleRangeBounds() {
	fmt.Println(xlsxtra.RangeBounds("A1:E6"))
	fmt.Println(xlsxtra.RangeBounds("$A$1:$E$6"))
	// invalid: no column name given
	fmt.Println(xlsxtra.RangeBounds("11:E6"))
	// invalid: row zero does not exist
	fmt.Println(xlsxtra.RangeBounds("A0:E6"))
	// Output:
	// 1 1 5 6 <nil>
	// 1 1 5 6 <nil>
	// 0 0 0 0 Invalid range "11:E6"
	// 0 0 0 0 Invalid range "A0:E6"
}
