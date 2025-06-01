package excelrecreator

import (
	"encoding/json"
	"fmt"
	"os"
	"strconv"
	"time"

	"github.com/prongbang/excelmetadata"
	"github.com/xuri/excelize/v2"
)

// Recreator handles the recreation of Excel files from metadata
type Recreator struct {
	file     *excelize.File
	metadata *excelmetadata.Metadata
	options  *Options
	styleMap map[int]int // Maps old style IDs to new style IDs
}

// Options configures the recreation behavior
type Options struct {
	PreserveFormulas       bool
	PreserveStyles         bool
	PreserveDataValidation bool
	SkipEmptyCells         bool
	DefaultSheetName       string
}

// DefaultOptions returns recommended default options
func DefaultOptions() *Options {
	return &Options{
		PreserveFormulas:       true,
		PreserveStyles:         true,
		PreserveDataValidation: true,
		SkipEmptyCells:         true,
		DefaultSheetName:       "Sheet",
	}
}

// New creates a new Recreator instance from metadata
func New(metadata *excelmetadata.Metadata, options *Options) *Recreator {
	if options == nil {
		options = DefaultOptions()
	}

	return &Recreator{
		file:     excelize.NewFile(),
		metadata: metadata,
		options:  options,
		styleMap: make(map[int]int),
	}
}

// NewFromJSON creates a new Recreator from JSON metadata
func NewFromJSON(jsonData []byte, options *Options) (*Recreator, error) {
	var metadata excelmetadata.Metadata
	if err := json.Unmarshal(jsonData, &metadata); err != nil {
		return nil, fmt.Errorf("failed to unmarshal JSON: %w", err)
	}

	return New(&metadata, options), nil
}

// NewFromJSONFile creates a new Recreator from a JSON file
func NewFromJSONFile(jsonPath string, options *Options) (*Recreator, error) {
	data, err := os.ReadFile(jsonPath)
	if err != nil {
		return nil, fmt.Errorf("failed to read JSON file: %w", err)
	}

	return NewFromJSON(data, options)
}

// Recreate performs the Excel file recreation
func (r *Recreator) Recreate() error {
	// Set document properties
	if err := r.recreateDocumentProperties(); err != nil {
		return fmt.Errorf("failed to recreate document properties: %w", err)
	}

	// Recreate styles first (to get style mapping)
	if r.options.PreserveStyles && len(r.metadata.Styles) > 0 {
		if err := r.recreateStyles(); err != nil {
			return fmt.Errorf("failed to recreate styles: %w", err)
		}
	}

	// Delete default sheet if we have sheets to create
	if len(r.metadata.Sheets) > 0 {
		r.file.DeleteSheet("Sheet1")
	}

	// Recreate each sheet
	for _, sheetMeta := range r.metadata.Sheets {
		if err := r.recreateSheet(sheetMeta); err != nil {
			return fmt.Errorf("failed to recreate sheet %s: %w", sheetMeta.Name, err)
		}
	}

	// Recreate defined names
	if r.options.PreserveFormulas && len(r.metadata.DefinedNames) > 0 {
		if err := r.recreateDefinedNames(); err != nil {
			return fmt.Errorf("failed to recreate defined names: %w", err)
		}
	}

	// Set active sheet to the first visible sheet
	for _, sheet := range r.metadata.Sheets {
		if sheet.Visible {
			r.file.SetActiveSheet(sheet.Index)
			break
		}
	}

	return nil
}

// Save saves the recreated Excel file
func (r *Recreator) Save(filename string) error {
	return r.file.SaveAs(filename)
}

// GetFile returns the underlying excelize.File for advanced operations
func (r *Recreator) GetFile() *excelize.File {
	return r.file
}

// Private recreation methods

func (r *Recreator) recreateDocumentProperties() error {
	props := &excelize.DocProperties{
		Title:          r.metadata.Properties.Title,
		Subject:        r.metadata.Properties.Subject,
		Creator:        r.metadata.Properties.Creator,
		Keywords:       r.metadata.Properties.Keywords,
		Description:    r.metadata.Properties.Description,
		LastModifiedBy: r.metadata.Properties.LastModifiedBy,
		Category:       r.metadata.Properties.Category,
		Version:        r.metadata.Properties.Version,
		Created:        r.metadata.Properties.Created,
		Modified:       r.metadata.Properties.Modified,
	}

	return r.file.SetDocProps(props)
}

func (r *Recreator) recreateStyles() error {
	for oldID, styleMeta := range r.metadata.Styles {
		style := &excelize.Style{}

		// Recreate font
		if styleMeta.Font != nil {
			style.Font = &excelize.Font{
				Bold:      styleMeta.Font.Bold,
				Italic:    styleMeta.Font.Italic,
				Underline: styleMeta.Font.Underline,
				Strike:    styleMeta.Font.Strike,
				Family:    styleMeta.Font.Family,
				Size:      styleMeta.Font.Size,
				Color:     styleMeta.Font.Color,
			}
		}

		// Recreate fill
		if styleMeta.Fill != nil && len(styleMeta.Fill.Color) > 0 {
			style.Fill = excelize.Fill{
				Type:    styleMeta.Fill.Type,
				Pattern: styleMeta.Fill.Pattern,
				Color:   styleMeta.Fill.Color,
			}
		}

		// Recreate borders
		if len(styleMeta.Border) > 0 {
			style.Border = []excelize.Border{}
			for _, borderMeta := range styleMeta.Border {
				style.Border = append(style.Border, excelize.Border{
					Type:  borderMeta.Type,
					Color: borderMeta.Color,
					Style: borderMeta.Style,
				})
			}
		}

		// Recreate alignment
		if styleMeta.Alignment != nil {
			style.Alignment = &excelize.Alignment{
				Horizontal:   styleMeta.Alignment.Horizontal,
				Vertical:     styleMeta.Alignment.Vertical,
				WrapText:     styleMeta.Alignment.WrapText,
				TextRotation: styleMeta.Alignment.TextRotation,
				Indent:       styleMeta.Alignment.Indent,
				ShrinkToFit:  styleMeta.Alignment.ShrinkToFit,
			}
		}

		// Recreate number format
		if styleMeta.NumberFormat != 0 {
			style.NumFmt = styleMeta.NumberFormat
		}

		// Recreate protection
		if styleMeta.Protection != nil {
			style.Protection = &excelize.Protection{
				Hidden: styleMeta.Protection.Hidden,
				Locked: styleMeta.Protection.Locked,
			}
		}

		// Create the style and map old ID to new ID
		newID, err := r.file.NewStyle(style)
		if err == nil {
			r.styleMap[oldID] = newID
		}
	}

	return nil
}

func (r *Recreator) recreateSheet(sheetMeta excelmetadata.SheetMetadata) error {
	sheetName := sheetMeta.Name
	if sheetName == "" {
		sheetName = fmt.Sprintf("%s%d", r.options.DefaultSheetName, sheetMeta.Index+1)
	}

	// Create sheet
	_, err := r.file.NewSheet(sheetName)
	if err != nil {
		return err
	}

	// Set visibility
	r.file.SetSheetVisible(sheetName, sheetMeta.Visible)

	// Set column widths
	for col, width := range sheetMeta.ColWidths {
		r.file.SetColWidth(sheetName, col, col, width)
	}

	// Set row heights
	for row, height := range sheetMeta.RowHeights {
		r.file.SetRowHeight(sheetName, row, height)
	}

	// Recreate cells
	if err := r.recreateCells(sheetName, sheetMeta.Cells); err != nil {
		return err
	}

	// Recreate merged cells
	for _, merge := range sheetMeta.MergedCells {
		r.file.MergeCell(sheetName, merge.StartCell, merge.EndCell)
	}

	// Recreate data validations
	if r.options.PreserveDataValidation {
		for _, dv := range sheetMeta.DataValidations {
			r.recreateDataValidation(sheetName, dv)
		}
	}

	// Recreate sheet protection
	if sheetMeta.Protection != nil && sheetMeta.Protection.Protected {
		r.recreateSheetProtection(sheetName, sheetMeta.Protection)
	}

	return nil
}

func (r *Recreator) recreateCells(sheetName string, cells []excelmetadata.CellMetadata) error {
	for _, cell := range cells {
		// Skip empty cells if option is set
		if r.options.SkipEmptyCells && cell.Value == nil && cell.Formula == "" {
			continue
		}

		// Set cell value or formula
		if cell.Formula != "" && r.options.PreserveFormulas {
			if err := r.file.SetCellFormula(sheetName, cell.Address, cell.Formula); err != nil {
				return err
			}
		} else if cell.Value != nil {
			// Handle different value types
			switch v := cell.Value.(type) {
			case float32:
				r.file.SetCellFloat(sheetName, cell.Address, float64(v), -1, 64)
			case float64:
				r.file.SetCellFloat(sheetName, cell.Address, v, -1, 64)
			case int:
				r.file.SetCellInt(sheetName, cell.Address, int64(v))
			case int8:
				r.file.SetCellInt(sheetName, cell.Address, int64(v))
			case int16:
				r.file.SetCellInt(sheetName, cell.Address, int64(v))
			case int32:
				r.file.SetCellInt(sheetName, cell.Address, int64(v))
			case bool:
				r.file.SetCellBool(sheetName, cell.Address, v)
			case time.Time:
				r.file.SetCellValue(sheetName, cell.Address, v)
			default:
				// Convert to string
				strVal := fmt.Sprintf("%v", v)
				// Try to parse as number
				if floatVal, err := strconv.ParseFloat(strVal, 64); err == nil {
					r.file.SetCellFloat(sheetName, cell.Address, floatVal, -1, 64)
				} else {
					r.file.SetCellValue(sheetName, cell.Address, strVal)
				}
			}
		}

		// Apply style
		if r.options.PreserveStyles && cell.StyleID != 0 {
			if newStyleID, exists := r.styleMap[cell.StyleID]; exists {
				r.file.SetCellStyle(sheetName, cell.Address, cell.Address, newStyleID)
			}
		}

		// Set hyperlink
		if cell.Hyperlink != nil {
			r.file.SetCellHyperLink(sheetName, cell.Address, cell.Hyperlink.Link, "Location")
		}
	}

	return nil
}

func (r *Recreator) recreateDataValidation(sheetName string, dv excelmetadata.DataValidation) error {
	validation := &excelize.DataValidation{
		Type:             dv.Type,
		Operator:         dv.Operator,
		Formula1:         dv.Formula1,
		Formula2:         dv.Formula2,
		ShowErrorMessage: dv.ShowError,
		ErrorTitle:       dv.ErrorTitle,
		Error:            dv.ErrorMessage,
		Sqref:            dv.Range,
	}

	return r.file.AddDataValidation(sheetName, validation)
}

func (r *Recreator) recreateSheetProtection(sheetName string, protection *excelmetadata.SheetProtection) error {
	editObjects := protection.EditObjects
	editScenarios := protection.EditScenarios
	selectLockedCells := protection.SelectLockedCells
	selectUnlockedCells := protection.SelectUnlockedCells

	opts := &excelize.SheetProtectionOptions{
		Password:            protection.Password,
		EditObjects:         editObjects,
		EditScenarios:       editScenarios,
		SelectLockedCells:   selectLockedCells,
		SelectUnlockedCells: selectUnlockedCells,
	}

	return r.file.ProtectSheet(sheetName, opts)
}

func (r *Recreator) recreateDefinedNames() error {
	for _, name := range r.metadata.DefinedNames {
		if err := r.file.SetDefinedName(&excelize.DefinedName{
			Name:     name.Name,
			RefersTo: name.RefersTo,
			Scope:    name.Scope,
		}); err != nil {
			return err
		}
	}
	return nil
}

// Utility functions

// QuickRecreate recreates an Excel file from metadata with default options
func QuickRecreate(metadata *excelmetadata.Metadata, outputPath string) error {
	recreator := New(metadata, DefaultOptions())

	if err := recreator.Recreate(); err != nil {
		return err
	}

	return recreator.Save(outputPath)
}

// QuickRecreateFromJSON recreates an Excel file from JSON metadata
func QuickRecreateFromJSON(jsonPath, outputPath string) error {
	recreator, err := NewFromJSONFile(jsonPath, DefaultOptions())
	if err != nil {
		return err
	}

	if err := recreator.Recreate(); err != nil {
		return err
	}

	return recreator.Save(outputPath)
}

// RecreateWithOptions recreates with custom options
func RecreateWithOptions(jsonPath, outputPath string, options *Options) error {
	recreator, err := NewFromJSONFile(jsonPath, options)
	if err != nil {
		return err
	}

	if err := recreator.Recreate(); err != nil {
		return err
	}

	return recreator.Save(outputPath)
}

// ValidateMetadata checks if metadata is valid for recreation
func ValidateMetadata(metadata *excelmetadata.Metadata) []string {
	var issues []string

	if metadata == nil {
		return []string{"metadata is nil"}
	}

	if len(metadata.Sheets) == 0 {
		issues = append(issues, "no sheets found in metadata")
	}

	for i, sheet := range metadata.Sheets {
		if sheet.Name == "" {
			issues = append(issues, fmt.Sprintf("sheet %d has no name", i))
		}

		// Check for invalid cell addresses
		for _, cell := range sheet.Cells {
			if _, _, err := excelize.CellNameToCoordinates(cell.Address); err != nil {
				issues = append(issues, fmt.Sprintf("invalid cell address: %s", cell.Address))
			}
		}

		// Check merged cells
		for _, merge := range sheet.MergedCells {
			if _, _, err := excelize.CellNameToCoordinates(merge.StartCell); err != nil {
				issues = append(issues, fmt.Sprintf("invalid merge start cell: %s", merge.StartCell))
			}
			if _, _, err := excelize.CellNameToCoordinates(merge.EndCell); err != nil {
				issues = append(issues, fmt.Sprintf("invalid merge end cell: %s", merge.EndCell))
			}
		}
	}

	return issues
}
