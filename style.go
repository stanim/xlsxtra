package xlsxtra

import "github.com/tealeg/xlsx"

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

// SetRowStyle set style to all cells of a row
func SetRowStyle(row *xlsx.Row, style *xlsx.Style) {
	for _, cell := range row.Cells {
		cell.SetStyle(style)
	}
}
