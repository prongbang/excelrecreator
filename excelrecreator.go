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
	File     *excelize.File
	Metadata *excelmetadata.Metadata
	Options  *Options
	StyleMap map[int]int // Maps old style IDs to new style IDs
}

// Options configures the recreation behavior
type Options struct {
	PreserveFormulas       bool
	PreserveStyles         bool
	PreserveDataValidation bool
	PreserveImages         bool
	SkipEmptyCells         bool
	DefaultSheetName       string
}

// DefaultOptions returns recommended default options
func DefaultOptions() *Options {
	return &Options{
		PreserveFormulas:       true,
		PreserveStyles:         true,
		PreserveDataValidation: true,
		PreserveImages:         true,
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
		File:     excelize.NewFile(),
		Metadata: metadata,
		Options:  options,
		StyleMap: make(map[int]int),
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
	if r.Options.PreserveStyles && len(r.Metadata.Styles) > 0 {
		if err := r.recreateStyles(); err != nil {
			return fmt.Errorf("failed to recreate styles: %w", err)
		}
	}

	// Recreate each sheet
	for _, sheetMeta := range r.Metadata.Sheets {
		if err := r.recreateSheet(sheetMeta); err != nil {
			return fmt.Errorf("failed to recreate sheet %s: %w", sheetMeta.Name, err)
		}
	}

	// Recreate defined names
	if r.Options.PreserveFormulas && len(r.Metadata.DefinedNames) > 0 {
		if err := r.recreateDefinedNames(); err != nil {
			return fmt.Errorf("failed to recreate defined names: %w", err)
		}
	}

	// Set active sheet to the first visible sheet
	for _, sheet := range r.Metadata.Sheets {
		if sheet.Visible {
			r.File.SetActiveSheet(sheet.Index)
			break
		}
	}

	return nil
}

// Save saves the recreated Excel file
func (r *Recreator) Save(filename string) error {
	return r.File.SaveAs(filename)
}

// GetFile returns the underlying excelize.File for advanced operations
func (r *Recreator) GetFile() *excelize.File {
	return r.File
}

// Private recreation methods

func (r *Recreator) recreateDocumentProperties() error {
	props := &excelize.DocProperties{
		Title:          r.Metadata.Properties.Title,
		Subject:        r.Metadata.Properties.Subject,
		Creator:        r.Metadata.Properties.Creator,
		Keywords:       r.Metadata.Properties.Keywords,
		Description:    r.Metadata.Properties.Description,
		LastModifiedBy: r.Metadata.Properties.LastModifiedBy,
		Category:       r.Metadata.Properties.Category,
		Version:        r.Metadata.Properties.Version,
		Created:        r.Metadata.Properties.Created,
		Modified:       r.Metadata.Properties.Modified,
	}

	return r.File.SetDocProps(props)
}

func (r *Recreator) recreateStyles() error {
	for oldID, styleMeta := range r.Metadata.Styles {
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
		newID, err := r.File.NewStyle(style)
		if err == nil {
			r.StyleMap[oldID] = newID
		}
	}

	return nil
}

func (r *Recreator) recreateSheet(sheetMeta excelmetadata.SheetMetadata) error {
	sheetName := sheetMeta.Name
	if sheetName == "" {
		sheetName = fmt.Sprintf("%s%d", r.Options.DefaultSheetName, sheetMeta.Index+1)
	}

	// Create sheet
	index, err := r.File.NewSheet(sheetName)
	if err != nil {
		return err
	}

	// Delete default sheet if we have sheets to create
	if len(r.Metadata.Sheets) > 0 {
		if sheetMeta.Index == 0 {
			r.File.SetActiveSheet(index)
			_ = r.File.DeleteSheet("Sheet1")
		}
	}

	// Set visibility
	r.File.SetSheetVisible(sheetName, sheetMeta.Visible)

	// Set column widths
	for col, width := range sheetMeta.ColWidths {
		r.File.SetColWidth(sheetName, col, col, width)
	}

	// Set row heights
	for row, height := range sheetMeta.RowHeights {
		r.File.SetRowHeight(sheetName, row, height)
	}

	// Recreate cells
	if err := r.recreateCells(sheetName, sheetMeta.Cells); err != nil {
		return err
	}

	// Recreate merged cells
	for _, merge := range sheetMeta.MergedCells {
		r.File.MergeCell(sheetName, merge.StartCell, merge.EndCell)
	}

	// Recreate data validations
	if r.Options.PreserveDataValidation {
		for _, dv := range sheetMeta.DataValidations {
			r.recreateDataValidation(sheetName, dv)
		}
	}

	// Recreate images
	if r.Options.PreserveImages {
		for _, img := range sheetMeta.Images {
			r.recreateImage(sheetName, &img)
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
		if r.Options.SkipEmptyCells && cell.Value == nil && cell.Formula == "" {
			continue
		}

		// Set cell value or formula
		if cell.Formula != "" && r.Options.PreserveFormulas {
			if err := r.File.SetCellFormula(sheetName, cell.Address, cell.Formula); err != nil {
				return err
			}
		} else if cell.Value != nil {
			// Handle different value types
			switch v := cell.Value.(type) {
			case float32:
				r.File.SetCellFloat(sheetName, cell.Address, float64(v), -1, 64)
			case float64:
				r.File.SetCellFloat(sheetName, cell.Address, v, -1, 64)
			case int:
				r.File.SetCellInt(sheetName, cell.Address, int64(v))
			case int8:
				r.File.SetCellInt(sheetName, cell.Address, int64(v))
			case int16:
				r.File.SetCellInt(sheetName, cell.Address, int64(v))
			case int32:
				r.File.SetCellInt(sheetName, cell.Address, int64(v))
			case bool:
				r.File.SetCellBool(sheetName, cell.Address, v)
			case time.Time:
				r.File.SetCellValue(sheetName, cell.Address, v)
			default:
				// Convert to string
				strVal := fmt.Sprintf("%v", v)
				// Try to parse as number
				if floatVal, err := strconv.ParseFloat(strVal, 64); err == nil {
					r.File.SetCellFloat(sheetName, cell.Address, floatVal, -1, 64)
				} else {
					r.File.SetCellValue(sheetName, cell.Address, strVal)
				}
			}
		}

		// Apply style
		if r.Options.PreserveStyles && cell.StyleID != 0 {
			if newStyleID, exists := r.StyleMap[cell.StyleID]; exists {
				r.File.SetCellStyle(sheetName, cell.Address, cell.Address, newStyleID)
			}
		}

		// Set hyperlink
		if cell.Hyperlink != nil {
			r.File.SetCellHyperLink(sheetName, cell.Address, cell.Hyperlink.Link, "Location")
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

	return r.File.AddDataValidation(sheetName, validation)
}

func (r *Recreator) recreateImage(sheetName string, img *excelmetadata.ImageMetadata) error {
	picture := &excelize.Picture{
		Extension: img.Extension,
		File:      img.File,
		Format: &excelize.GraphicOptions{
			AltText:             img.Format.AltText,
			PrintObject:         img.Format.PrintObject,
			Locked:              img.Format.Locked,
			LockAspectRatio:     img.Format.LockAspectRatio,
			AutoFit:             img.Format.AutoFit,
			AutoFitIgnoreAspect: img.Format.AutoFitIgnoreAspect,
			OffsetX:             img.Format.OffsetX,
			OffsetY:             img.Format.OffsetY,
			ScaleX:              img.Format.ScaleX,
			ScaleY:              img.Format.ScaleY,
			Hyperlink:           img.Format.Hyperlink,
			HyperlinkType:       img.Format.HyperlinkType,
			Positioning:         img.Format.Positioning,
		},
		InsertType: excelize.PictureInsertType(img.InsertType),
	}

	return r.File.AddPictureFromBytes(sheetName, img.Cell, picture)
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

	return r.File.ProtectSheet(sheetName, opts)
}

func (r *Recreator) recreateDefinedNames() error {
	for _, name := range r.Metadata.DefinedNames {
		if err := r.File.SetDefinedName(&excelize.DefinedName{
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
