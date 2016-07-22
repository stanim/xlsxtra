// Package xlsxtra provides extra utilities for the xlsx
// package (https://github.com/tealeg/xlsx) to manipulate
// excel files:
//
// - Sort(): multi-column sort of selected rows
//
// - AddBool(), AddInt(), AddFloat(), ...: shortcut to add
// a cell to a row with the right type.
//
// - NewStyle(): create a style and set the ApplyFill,
// ApplyFont, ApplyBorder and ApplyAlignment automatically.
//
// - NewStyles(): create a slice of styles based on a color
// palette
//
// - Sheets: access sheets by name instead of by index
//
// - Col: access cell values of a row by column header title
//
// - SetRowStyle: set style of all cells in a row
//
// - ToString: convert a xlsx.Row to a slice of strings
//
// See Col(umn) and Sort example for a quick introduction.
package xlsxtra
