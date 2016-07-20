package xlsxtra

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

// Default fonts and styles
var (
	BoldFont = &xlsx.Font{
		Size: 10, Name: "Arial", Bold: true}
	BoldStyle   = NewStyle("", BoldFont, nil, nil)
	HeaderStyle = NewStyle("00CCCCCC", BoldFont, nil, nil)
	WrapText    = &xlsx.Alignment{
		Horizontal:   "general",
		Vertical:     "bottom",
		Indent:       0,
		ShrinkToFit:  false,
		TextRotation: 0,
		WrapText:     true}
	WrapStyle = NewStyle("", nil, nil, WrapText)
	// http://www.color-hex.com/color-palette/20772
	Palette = []string{
		"00ffaeae",
		"00ffc49c",
		"00fffb96",
		"00caff7e",
		"009bf9bd",
	}
	PaletteStyles = NewStyles(Palette, BoldFont,
		nil, nil)
)

// Tabs retrieves a sheet by label
type Tabs struct {
	sheets   map[string]*xlsx.Sheet
	filename string
}

// NewTabs creates new tabs from sheet names of an already
// opened xlsx.File
func NewTabs(file *xlsx.File, fn string) *Tabs {
	tabs := Tabs{
		sheets:   make(map[string]*xlsx.Sheet),
		filename: fn,
	}
	for _, sheet := range file.Sheets {
		tabs.sheets[sheet.Name] = sheet
	}
	return &tabs
}

// OpenTabs opens an xlsx file from disk and returns its
// tabs
func OpenTabs(fn string) (*Tabs, error) {
	xlFile, err := xlsx.OpenFile(fn)
	if err != nil {
		return nil, err
	}
	return NewTabs(xlFile, fn), nil
}

// Get returns a certain tab sheet
func (t Tabs) Get(name string) (*xlsx.Sheet, error) {
	sheet, ok := t.sheets[name]
	if !ok {
		return nil, fmt.Errorf(
			"file %q does not contain sheet %q",
			t.filename, name)
	}
	return sheet, nil
}

// OpenSheet open a sheet from an xlsx file. If you need
// to use multiple sheets from one file use the Tabs type
// instead.
func OpenSheet(fn string, name string) (
	*xlsx.Sheet, error) {
	tabs, err := OpenTabs(fn)
	if err != nil {
		return nil, err
	}
	return tabs.Get(name)
}

// Col retrieves values by header label from a row
type Col map[string]int

// NewCol creates a new Col from a header row
func NewCol(row *xlsx.Row) Col {
	col := make(Col)
	for i, cell := range row.Cells {
		if cell.Value != "" {
			col[cell.Value] = i
		}
	}
	return col
}

func (c Col) index(
	row *xlsx.Row, header string) (int, error) {
	if i, ok := c[header]; ok {
		if i < len(row.Cells) {
			return i, nil
		}
		return i, fmt.Errorf(
			"Index %d out of range for column header: %s (first column: %s)\nTry to set a value in a column to the far right.",
			i, header, row.Cells[0].Value)
	}
	return 0, fmt.Errorf("Unknown column header: %s (%#v)",
		header, c)
}

// Bool value of (row,col) in spreadsheet
func (c Col) Bool(row *xlsx.Row, header string) (bool,
	error) {
	val, err := c.String(row, header)
	if err != nil {
		return false, err
	}
	val = strings.ToLower(val)
	return val == "ja" || val == "yes", nil
}

// BoolMap value of (row,col) in spreadsheet
func (c Col) BoolMap(row *xlsx.Row, headers []string) (
	map[string]bool, error) {
	var err error
	bmap := make(map[string]bool)
	for _, header := range headers {
		bmap[header], err = c.Bool(row, header)
		if err != nil {
			return nil, err
		}
	}
	return bmap, nil
}

// Int value of (row,col) in spreadsheet
func (c Col) Int(row *xlsx.Row, header string) (int,
	error) {
	i, err := c.index(row, header)
	if err != nil {
		return 0, err
	}
	return row.Cells[i].Int()
}

// Float value of (row,col) in spreadsheet
func (c Col) Float(row *xlsx.Row, header string) (float64,
	error) {
	i, err := c.index(row, header)
	if err != nil {
		return 0, err
	}
	val := row.Cells[i].Value
	if strings.HasPrefix(val, "€") {
		val = strings.TrimLeft(val, "€ ")
	}
	f, err := strconv.ParseFloat(val, 64)
	if err != nil {
		return math.NaN(), err
	}
	return f, nil
}

// String value of (row,col) in spreadsheet
func (c Col) String(row *xlsx.Row, header string) (string,
	error) {
	i, err := c.index(row, header)
	if err != nil {
		return "", err
	}
	if i >= len(row.Cells) {
		return "", nil
	}
	return row.Cells[i].String()
}

// AddBool adds a cell with bool as 1 or 0 to a row
func AddBool(row *xlsx.Row, x bool) *xlsx.Cell {
	var cell *xlsx.Cell
	if x {
		cell = AddInt(row, 1)
	} else {
		cell = AddInt(row, 0)
	}
	return cell
}

// AddFloat adds a cell with float64 value to a row
func AddFloat(row *xlsx.Row, x float64,
	format string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetFloatWithFormat(x, format)
	return cell
}

// AddFormula adds a cell with formula to a row
func AddFormula(row *xlsx.Row, formula string,
	format string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetFormula(formula)
	cell.NumFmt = format
	return cell
}

// AddInt adds a cell with int value to a row
func AddInt(row *xlsx.Row, x int) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetInt(x)
	return cell
}

// AddString adds a cell with string value to a row
func AddString(row *xlsx.Row, x string) *xlsx.Cell {
	cell := row.AddCell()
	cell.SetString(x)
	return cell
}

// StringFloatMap converts column with days string into
// a map of floats.
func (c Col) StringFloatMap(row *xlsx.Row, header string,
	dmap map[string]float64, val float64, sep string,
	chars int) error {
	// days
	s, err := c.String(row, header)
	if err != nil {
		return err
	}
	ds := strings.Split(s, sep)
	for _, d := range ds {
		if chars > 0 && len(s) > chars {
			d = d[:chars]
		}
		dmap[d] = val
	}
	return nil
}

// NewStyle creates a new style with color and boldness
func NewStyle(color string, font *xlsx.Font,
	border *xlsx.Border, align *xlsx.Alignment) *xlsx.Style {
	style := xlsx.NewStyle()
	if color != "" {
		style.Fill = *xlsx.NewFill("solid", color, color)
		style.ApplyFill = true
	} else {
		style.Fill = *xlsx.DefaultFill()
	}
	if font != nil {
		style.Font = *font
		style.ApplyFont = true
	} else {
		style.Font = *xlsx.DefaultFont()
	}
	if border != nil {
		style.Border = *border
		style.ApplyBorder = true
	} else {
		style.Border = *xlsx.DefaultBorder()
	}
	if align != nil {
		style.Alignment = *align
		style.ApplyAlignment = true
	} else {
		style.Alignment = *xlsx.DefaultAlignment()
	}
	return style
}

// NewStyles creates styles with color and boldness
func NewStyles(colors []string, font *xlsx.Font,
	border *xlsx.Border,
	align *xlsx.Alignment) []*xlsx.Style {
	styles := make([]*xlsx.Style, len(colors))
	for i, color := range colors {
		styles[i] = NewStyle(color, font, border, align)
	}
	return styles
}

// SetStyleRow set style to all cells of a row
func SetStyleRow(row *xlsx.Row, style *xlsx.Style) {
	for _, cell := range row.Cells {
		cell.SetStyle(style)
	}
}
