package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/prongbang/excelmetadata"
	"github.com/prongbang/excelrecreator"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Example 1: Basic recreation from JSON file
	basicExample()

	// Example 2: Recreation with custom options
	customOptionsExample()

	// Example 3: Modify metadata before recreation
	modifyMetadataExample()

	// Example 4: Partial recreation
	partialRecreationExample()

	// Example 5: Batch processing
	batchProcessingExample()

	// Example 6: Template creation
	templateCreationExample()
}

// Example 1: Basic recreation from JSON file
func basicExample() {
	fmt.Println("=== Basic Recreation Example ===")

	// Method 1: Direct recreation from JSON file
	err := excelrecreator.QuickRecreateFromJSON("metadata.json", "recreated.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Successfully recreated Excel file: recreated.xlsx")

	// Method 2: Load metadata first
	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	err = excelrecreator.QuickRecreate(metadata, "recreated2.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Successfully recreated Excel file: recreated2.xlsx")
}

// Example 2: Recreation with custom options
func customOptionsExample() {
	fmt.Println("\n=== Custom Options Example ===")

	// Configure what to preserve
	options := &excelrecreator.Options{
		PreserveFormulas:       true,
		PreserveStyles:         true,
		PreserveDataValidation: true,
		SkipEmptyCells:         true,
		DefaultSheetName:       "Data",
	}

	// Recreate with custom options
	err := excelrecreator.RecreateWithOptions("metadata.json", "custom_recreated.xlsx", options)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Recreated with custom options: custom_recreated.xlsx")
	fmt.Println("- Protection removed")
	fmt.Println("- Empty cells skipped")
}

// Example 3: Modify metadata before recreation
func modifyMetadataExample() {
	fmt.Println("\n=== Modify Metadata Example ===")

	// Load metadata
	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Modify metadata before recreation
	// 1. Update document properties
	metadata.Properties.Title = "Modified Excel File"
	metadata.Properties.Creator = "Excel Recreator"
	metadata.Properties.Description = "Recreated from metadata with modifications"

	// 2. Remove specific sheets
	var filteredSheets []excelmetadata.SheetMetadata
	for _, sheet := range metadata.Sheets {
		if sheet.Name != "HiddenSheet" { // Skip hidden sheet
			filteredSheets = append(filteredSheets, sheet)
		}
	}
	metadata.Sheets = filteredSheets

	// 3. Add prefix to all cell values
	for i, sheet := range metadata.Sheets {
		for j, cell := range sheet.Cells {
			if cell.Formula == "" && cell.Value != nil {
				// Add prefix to text values
				if strVal, ok := cell.Value.(string); ok {
					metadata.Sheets[i].Cells[j].Value = "MOD: " + strVal
				}
			}
		}
	}

	// 4. Remove all hyperlinks
	for i, sheet := range metadata.Sheets {
		for j := range sheet.Cells {
			metadata.Sheets[i].Cells[j].Hyperlink = nil
		}
	}

	// Recreate modified Excel
	recreator := excelrecreator.New(metadata, nil)
	if err := recreator.Recreate(); err != nil {
		log.Fatal(err)
	}

	if err := recreator.Save("modified_recreated.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created modified Excel file: modified_recreated.xlsx")
	fmt.Println("- Updated properties")
	fmt.Println("- Removed hidden sheets")
	fmt.Println("- Added prefix to cell values")
	fmt.Println("- Removed hyperlinks")
}

// Example 4: Partial recreation (specific sheets only)
func partialRecreationExample() {
	fmt.Println("\n=== Partial Recreation Example ===")

	// Load metadata
	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Create new metadata with only specific sheets
	partialMetadata := &excelmetadata.Metadata{
		Filename:   metadata.Filename + "_partial",
		Properties: metadata.Properties,
		Styles:     metadata.Styles, // Keep all styles
		Sheets:     []excelmetadata.SheetMetadata{},
	}

	// Select only sheets that contain "Data" in the name
	for _, sheet := range metadata.Sheets {
		if contains(sheet.Name, "Data") {
			partialMetadata.Sheets = append(partialMetadata.Sheets, sheet)
			fmt.Printf("Including sheet: %s\n", sheet.Name)
		}
	}

	// Recreate partial Excel
	if err := excelrecreator.QuickRecreate(partialMetadata, "partial_recreated.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created partial Excel file: partial_recreated.xlsx")
}

// Example 5: Batch processing multiple files
func batchProcessingExample() {
	fmt.Println("\n=== Batch Processing Example ===")

	// List of JSON files to process
	jsonFiles := []string{
		"metadata1.json",
		"metadata2.json",
		"metadata3.json",
	}

	// Process each file
	for _, jsonFile := range jsonFiles {
		// Check if file exists
		if _, err := os.Stat(jsonFile); os.IsNotExist(err) {
			fmt.Printf("Skipping %s (not found)\n", jsonFile)
			continue
		}

		// Generate output filename
		outputFile := strings.Replace(jsonFile, ".json", "_recreated.xlsx", 1)

		// Recreate Excel
		err := excelrecreator.QuickRecreateFromJSON(jsonFile, outputFile)
		if err != nil {
			fmt.Printf("Error processing %s: %v\n", jsonFile, err)
			continue
		}

		fmt.Printf("Successfully processed: %s -> %s\n", jsonFile, outputFile)
	}
}

// Example 6: Create template from metadata
func templateCreationExample() {
	fmt.Println("\n=== Template Creation Example ===")

	// Create minimal metadata for a template
	templateMetadata := &excelmetadata.Metadata{
		Filename: "template.xlsx",
		Properties: excelmetadata.DocumentProperties{
			Title:       "Excel Template",
			Creator:     "Template Generator",
			Description: "Template for data entry",
		},
		Sheets: []excelmetadata.SheetMetadata{
			{
				Index:   0,
				Name:    "DataEntry",
				Visible: true,
				Dimensions: excelmetadata.SheetDimensions{
					StartCell: "A1",
					EndCell:   "E10",
					RowCount:  10,
					ColCount:  5,
				},
				ColWidths: map[string]float64{
					"A": 15,
					"B": 20,
					"C": 20,
					"D": 15,
					"E": 25,
				},
				Cells: []excelmetadata.CellMetadata{
					// Headers
					{Address: "A1", Value: "ID", StyleID: 1},
					{Address: "B1", Value: "Name", StyleID: 1},
					{Address: "C1", Value: "Email", StyleID: 1},
					{Address: "D1", Value: "Phone", StyleID: 1},
					{Address: "E1", Value: "Notes", StyleID: 1},
					// Sample data with formulas
					{Address: "A2", Value: 1},
					{Address: "B2", Value: "John Doe"},
					{Address: "C2", Value: "john@example.com"},
					{Address: "D2", Value: "123-456-7890"},
					{Address: "E2", Value: "Sample entry"},
					// Formula example
					{Address: "A10", Value: "Total:", StyleID: 1},
					{Address: "B10", Formula: "COUNTA(A2:A9)", StyleID: 1},
				},
				MergedCells: []excelmetadata.MergedCell{
					{StartCell: "A1", EndCell: "A1"}, // Headers can be merged if needed
				},
				DataValidations: []excelmetadata.DataValidation{
					{
						Range:        "C2:C9",
						Type:         "custom",
						Formula1:     `ISERROR(FIND(" ",C2))*(FIND("@",C2)>0)*(FIND(".",C2)>FIND("@",C2))`,
						ShowError:    true,
						ErrorTitle:   Pointer("Invalid Email"),
						ErrorMessage: Pointer("Please enter a valid email address"),
					},
				},
			},
			{
				Index:   1,
				Name:    "Instructions",
				Visible: true,
				Cells: []excelmetadata.CellMetadata{
					{Address: "A1", Value: "How to use this template:", StyleID: 1},
					{Address: "A3", Value: "1. Fill in the data starting from row 2"},
					{Address: "A4", Value: "2. Email validation is enabled for column C"},
					{Address: "A5", Value: "3. The total count is calculated automatically"},
					{Address: "A7", Value: "Tips:", StyleID: 1},
					{Address: "A8", Value: "- Use Tab to move between cells"},
					{Address: "A9", Value: "- Press Ctrl+S to save your work"},
				},
			},
		},
		Styles: map[int]excelmetadata.StyleDetails{
			1: { // Header style
				Font: &excelmetadata.FontStyle{
					Bold:  true,
					Size:  12,
					Color: "#000000",
				},
				Fill: &excelmetadata.FillStyle{
					Type:    "pattern",
					Pattern: 1,
					Color:   []string{"#E0E0E0"},
				},
				Alignment: &excelmetadata.AlignmentStyle{
					Horizontal: "center",
					Vertical:   "center",
				},
			},
		},
	}

	// Create the template
	recreator := excelrecreator.New(templateMetadata, nil)
	if err := recreator.Recreate(); err != nil {
		log.Fatal(err)
	}

	// Add additional formatting using the excelize API directly
	file := recreator.GetFile()

	// Add conditional formatting for the ID column
	style, _ := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#9A0511"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#FEC7CE"}, Pattern: 1},
	})
	opt := []excelize.ConditionalFormatOptions{
		{Type: "duplicate", Criteria: "=", Format: &style},
	}
	file.SetConditionalFormat("DataEntry", "A2:A9", opt)

	// Protect the sheet but allow editing in data cells
	file.ProtectSheet("DataEntry", &excelize.SheetProtectionOptions{
		Password:            "",
		SelectLockedCells:   true,
		SelectUnlockedCells: true,
	})

	// Save the template
	if err := recreator.Save("template.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created template file: template.xlsx")
	fmt.Println("- Data entry sheet with headers")
	fmt.Println("- Email validation on column C")
	fmt.Println("- Automatic row counting")
	fmt.Println("- Instructions sheet")
}

// Advanced Examples

// Example: Merge multiple metadata files into one Excel
func mergeMetadataExample() {
	fmt.Println("\n=== Merge Metadata Example ===")

	// Load multiple metadata files
	metadataFiles := []string{"sales_q1.json", "sales_q2.json", "sales_q3.json", "sales_q4.json"}

	// Create merged metadata
	mergedMetadata := &excelmetadata.Metadata{
		Filename: "sales_yearly.xlsx",
		Properties: excelmetadata.DocumentProperties{
			Title:   "Yearly Sales Report",
			Creator: "Sales Department",
		},
		Sheets: []excelmetadata.SheetMetadata{},
		Styles: map[int]excelmetadata.StyleDetails{},
	}

	// Merge each file
	for _, file := range metadataFiles {
		if metadata, err := loadMetadataFromFile(file); err == nil {
			// Add all sheets from this file
			mergedMetadata.Sheets = append(mergedMetadata.Sheets, metadata.Sheets...)

			// Merge styles (avoiding duplicates)
			for id, style := range metadata.Styles {
				if _, exists := mergedMetadata.Styles[id]; !exists {
					mergedMetadata.Styles[id] = style
				}
			}
		}
	}

	// Recreate merged Excel
	if err := excelrecreator.QuickRecreate(mergedMetadata, "sales_yearly.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Merged %d files into sales_yearly.xlsx\n", len(metadataFiles))
}

// Example: Clean and optimize metadata before recreation
func cleanMetadataExample() {
	fmt.Println("\n=== Clean Metadata Example ===")

	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Clean operations
	cleaned := cleanMetadata(metadata)

	// Recreate cleaned Excel
	if err := excelrecreator.QuickRecreate(cleaned, "cleaned.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created cleaned Excel file: cleaned.xlsx")
}

// Example: Validate metadata before recreation
func validateMetadataExample() {
	fmt.Println("\n=== Validate Metadata Example ===")

	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Validate metadata
	issues := excelrecreator.ValidateMetadata(metadata)

	if len(issues) > 0 {
		fmt.Println("Validation issues found:")
		for _, issue := range issues {
			fmt.Printf("  - %s\n", issue)
		}

		// Attempt to fix issues
		metadata = fixMetadataIssues(metadata, issues)
		fmt.Println("\nAttempted to fix issues")
	} else {
		fmt.Println("No validation issues found")
	}

	// Recreate Excel
	if err := excelrecreator.QuickRecreate(metadata, "validated.xlsx"); err != nil {
		log.Fatal(err)
	}
}

// Example: Transform data during recreation
func transformDataExample() {
	fmt.Println("\n=== Transform Data Example ===")

	metadata, err := loadMetadataFromFile("metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Apply transformations
	for i, sheet := range metadata.Sheets {
		for j, cell := range sheet.Cells {
			// Example: Convert all numbers to currency format
			if _, ok := cell.Value.(float64); ok {
				metadata.Sheets[i].Cells[j].StyleID = 2 // Assuming style 2 is currency
			}

			// Example: Upper case all text in column A
			if strings.HasPrefix(cell.Address, "A") {
				if strVal, ok := cell.Value.(string); ok {
					metadata.Sheets[i].Cells[j].Value = strings.ToUpper(strVal)
				}
			}
		}
	}

	// Recreate with transformations
	if err := excelrecreator.QuickRecreate(metadata, "transformed.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created transformed Excel file: transformed.xlsx")
}

// Helper functions

func loadMetadataFromFile(path string) (*excelmetadata.Metadata, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	var metadata excelmetadata.Metadata
	if err := json.Unmarshal(data, &metadata); err != nil {
		return nil, err
	}

	return &metadata, nil
}

func contains(s, substr string) bool {
	return strings.Contains(strings.ToLower(s), strings.ToLower(substr))
}

func cleanMetadata(metadata *excelmetadata.Metadata) *excelmetadata.Metadata {
	cleaned := *metadata // Copy

	// Remove empty cells
	for i, sheet := range cleaned.Sheets {
		var nonEmptyCells []excelmetadata.CellMetadata
		for _, cell := range sheet.Cells {
			if cell.Value != nil || cell.Formula != "" {
				nonEmptyCells = append(nonEmptyCells, cell)
			}
		}
		cleaned.Sheets[i].Cells = nonEmptyCells
	}

	// Remove unused styles
	usedStyles := make(map[int]bool)
	for _, sheet := range cleaned.Sheets {
		for _, cell := range sheet.Cells {
			usedStyles[cell.StyleID] = true
		}
	}

	newStyles := make(map[int]excelmetadata.StyleDetails)
	for id, style := range cleaned.Styles {
		if usedStyles[id] {
			newStyles[id] = style
		}
	}
	cleaned.Styles = newStyles

	return &cleaned
}

func fixMetadataIssues(metadata *excelmetadata.Metadata, issues []string) *excelmetadata.Metadata {
	fixed := *metadata // Copy

	for _, issue := range issues {
		if strings.Contains(issue, "has no name") {
			// Fix sheets with no names
			for i, sheet := range fixed.Sheets {
				if sheet.Name == "" {
					fixed.Sheets[i].Name = fmt.Sprintf("Sheet%d", i+1)
				}
			}
		}

		if strings.Contains(issue, "invalid cell address") {
			// Remove cells with invalid addresses
			for i, sheet := range fixed.Sheets {
				var validCells []excelmetadata.CellMetadata
				for _, cell := range sheet.Cells {
					if _, _, err := excelize.CellNameToCoordinates(cell.Address); err == nil {
						validCells = append(validCells, cell)
					}
				}
				fixed.Sheets[i].Cells = validCells
			}
		}
	}

	return &fixed
}

// Example: Progress monitoring for large files
func progressMonitoringExample() {
	fmt.Println("\n=== Progress Monitoring Example ===")

	metadata, err := loadMetadataFromFile("large_metadata.json")
	if err != nil {
		log.Fatal(err)
	}

	// Create recreator with progress callback
	recreator := excelrecreator.New(metadata, nil)

	totalSheets := len(metadata.Sheets)
	fmt.Printf("Processing %d sheets...\n", totalSheets)

	// Note: This is a conceptual example - the actual library would need
	// to support progress callbacks

	if err := recreator.Recreate(); err != nil {
		log.Fatal(err)
	}

	if err := recreator.Save("large_recreated.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("\nCompleted!")
}

// Example: Error handling and recovery
func errorHandlingExample() {
	fmt.Println("\n=== Error Handling Example ===")

	// Try to load potentially corrupted metadata
	metadata, err := loadMetadataFromFile("corrupted_metadata.json")
	if err != nil {
		fmt.Printf("Error loading metadata: %v\n", err)
		fmt.Println("Creating minimal fallback Excel...")

		// Create minimal fallback
		fallbackMetadata := &excelmetadata.Metadata{
			Filename: "fallback.xlsx",
			Properties: excelmetadata.DocumentProperties{
				Title: "Fallback Document",
			},
			Sheets: []excelmetadata.SheetMetadata{
				{
					Name: "Error",
					Cells: []excelmetadata.CellMetadata{
						{
							Address: "A1",
							Value:   "Error loading original metadata",
						},
						{
							Address: "A2",
							Value:   err.Error(),
						},
					},
				},
			},
		}

		if err := excelrecreator.QuickRecreate(fallbackMetadata, "fallback.xlsx"); err != nil {
			log.Fatal("Failed to create fallback:", err)
		}

		fmt.Println("Created fallback.xlsx")
		return
	}

	// Try recreation with error recovery
	options := &excelrecreator.Options{
		PreserveFormulas: false, // Disable formulas if they cause issues
		SkipEmptyCells:   true,
	}

	recreator := excelrecreator.New(metadata, options)

	if err := recreator.Recreate(); err != nil {
		fmt.Printf("Recreation error: %v\n", err)
		fmt.Println("Attempting partial recovery...")

		// Could implement partial recovery logic here
	}

	if err := recreator.Save("recovered.xlsx"); err != nil {
		log.Fatal("Failed to save:", err)
	}

	fmt.Println("Successfully handled errors and created recovered.xlsx")
}

func Pointer[T any](data T) *T {
	return &data
}
