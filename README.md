# ExcelRecreator

A Go library for recreating Excel files from metadata extracted by [excelmetadata](https://github.com/prongbang/excelmetadata). This library allows you to reconstruct complete Excel files (.xlsx) from JSON metadata, preserving formatting, formulas, styles, and structure.

## Features

- 📄 **Complete Excel Recreation**
  - Document properties (title, author, dates, etc.)
  - Multiple sheets with proper visibility settings
  - Cell values with type preservation
  - Formulas and calculations
  - Merged cells

- 🎨 **Style Preservation**
  - Font formatting (bold, italic, color, size)
  - Cell fills and patterns
  - Borders and alignment
  - Number formats
  - Cell protection settings

- 🔧 **Advanced Features**
  - Data validation rules
  - Sheet protection
  - Named ranges (defined names)
  - Hyperlinks
  - Custom row heights and column widths

- ⚙️ **Flexible Options**
  - Selective feature preservation
  - Empty cell handling
  - Custom default sheet names

## Installation

```bash
go get github.com/prongbang/excelrecreator
```

## Requirements

- Go 1.18 or higher
- github.com/xuri/excelize/v2 v2.9.1+
- github.com/prongbang/excelmetadata (for metadata structures)

## Quick Start

### Basic Usage

```go
package main

import (
    "log"
    "github.com/prongbang/excelmetadata"
    "github.com/prongbang/excelrecreator"
)

func main() {
    // Option 1: From metadata struct
    metadata, err := excelmetadata.QuickExtract("original.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    err = excelrecreator.QuickRecreate(metadata, "recreated.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    // Option 2: From JSON file
    err = excelrecreator.QuickRecreateFromJSON("metadata.json", "output.xlsx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Advanced Usage with Options

```go
// Configure recreation options
options := &excelrecreator.Options{
    PreserveFormulas:       true,   // Keep formulas
    PreserveStyles:         true,   // Keep all formatting
    PreserveDataValidation: true,   // Keep validation rules
    PreserveImages: 				true,   // Keep images
    SkipEmptyCells:         true,   // Don't create empty cells
    DefaultSheetName:       "Sheet", // Default name for unnamed sheets
}

// Create recreator with options
recreator, err := excelrecreator.NewFromJSONFile("metadata.json", options)
if err != nil {
    log.Fatal(err)
}

// Recreate the Excel file
if err := recreator.Recreate(); err != nil {
    log.Fatal(err)
}

// Save to file
if err := recreator.Save("recreated_with_options.xlsx"); err != nil {
    log.Fatal(err)
}
```

### Working with Metadata Directly

```go
// Load metadata from JSON
jsonData, _ := os.ReadFile("metadata.json")
recreator, err := excelrecreator.NewFromJSON(jsonData, nil)
if err != nil {
    log.Fatal(err)
}

// Or create from metadata struct
metadata := &excelmetadata.Metadata{
    // ... populate metadata
}
recreator = excelrecreator.New(metadata, nil)

// Access the underlying excelize.File for advanced operations
file := recreator.GetFile()
// ... perform additional operations

// Recreate and save
recreator.Recreate()
recreator.Save("output.xlsx")
```

## Recreation Options

| Option | Description | Default |
|--------|-------------|---------|
| `PreserveFormulas` | Recreate cell formulas | `true` |
| `PreserveStyles` | Apply all style formatting | `true` |
| `PreserveImages` | Apply all images | `true` |
| `PreserveDataValidation` | Apply data validation rules | `true` |
| `SkipEmptyCells` | Skip cells with no value or formula | `true` |
| `DefaultSheetName` | Base name for unnamed sheets | `"Sheet"` |

## Metadata Validation

Before recreating, you can validate the metadata:

```go
metadata, _ := excelmetadata.QuickExtract("source.xlsx")
issues := excelrecreator.ValidateMetadata(metadata)

if len(issues) > 0 {
    log.Println("Validation issues found:")
    for _, issue := range issues {
        log.Printf("- %s\n", issue)
    }
}
```

## Complete Workflow Example

```go
package main

import (
    "encoding/json"
    "fmt"
    "log"
    "os"

    "github.com/prongbang/excelmetadata"
    "github.com/prongbang/excelrecreator"
)

func main() {
    // Step 1: Extract metadata from original Excel
    fmt.Println("Extracting metadata...")
    metadata, err := excelmetadata.QuickExtract("original.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    // Step 2: Save metadata to JSON (optional)
    jsonData, _ := json.MarshalIndent(metadata, "", "  ")
    os.WriteFile("metadata.json", jsonData, 0644)

    // Step 3: Modify metadata if needed
    metadata.Properties.Title = "Modified Document"
    metadata.Properties.Creator = "ExcelRecreator"

    // Step 4: Recreate Excel from metadata
    fmt.Println("Recreating Excel file...")
    err = excelrecreator.QuickRecreate(metadata, "recreated.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("Success! Created recreated.xlsx")
}
```

## Supported Features

### ✅ Fully Supported
- Document properties
- Sheet structure and visibility
- Cell values (all types: string, number, boolean, date/time)
- Cell formulas
- Cell styles (font, fill, border, alignment, number format)
- Merged cells
- Row heights and column widths
- Data validation rules
- Sheet protection
- Named ranges
- Hyperlinks
- Images

### ⚠️ Limitations
- Charts and pivot tables (not implemented)
- VBA macros (not supported by excelize)
- Some advanced Excel features

## Error Handling

The library provides detailed error messages for debugging:

```go
recreator, err := excelrecreator.NewFromJSONFile("metadata.json", nil)
if err != nil {
    // Handle JSON reading/parsing errors
    log.Printf("Failed to load metadata: %v", err)
    return
}

if err := recreator.Recreate(); err != nil {
    // Handle recreation errors
    log.Printf("Failed to recreate Excel: %v", err)
    return
}

if err := recreator.Save("output.xlsx"); err != nil {
    // Handle file saving errors
    log.Printf("Failed to save file: %v", err)
    return
}
```

## Use Cases

1. **Excel File Recovery** - Recreate Excel files from metadata backups
2. **Template Generation** - Create new Excel files from template metadata
3. **Batch Processing** - Convert multiple JSON metadata files to Excel
4. **Data Migration** - Transfer Excel structures between systems
5. **Version Control** - Store Excel structures in Git as JSON

## Example: Batch Processing

```go
func batchConvert(jsonDir, outputDir string) error {
    files, err := os.ReadDir(jsonDir)
    if err != nil {
        return err
    }

    for _, file := range files {
        if filepath.Ext(file.Name()) != ".json" {
            continue
        }

        jsonPath := filepath.Join(jsonDir, file.Name())
        outputPath := filepath.Join(outputDir,
            strings.TrimSuffix(file.Name(), ".json") + ".xlsx")

        err := excelrecreator.QuickRecreateFromJSON(jsonPath, outputPath)
        if err != nil {
            log.Printf("Failed to process %s: %v", file.Name(), err)
            continue
        }

        fmt.Printf("Created: %s\n", outputPath)
    }

    return nil
}
```

## Example: Creating Excel from Scratch

```go
func createFromScratch() error {
    // Build metadata programmatically
    metadata := &excelmetadata.Metadata{
        Properties: excelmetadata.DocumentProperties{
            Title:   "Sales Report",
            Creator: "Sales System",
        },
        Sheets: []excelmetadata.SheetMetadata{
            {
                Name: "Q1 Sales",
                Cells: []excelmetadata.CellMetadata{
                    {Address: "A1", Value: "Product"},
                    {Address: "B1", Value: "Revenue"},
                    {Address: "A2", Value: "Widget A"},
                    {Address: "B2", Value: 1000.50},
                    {Address: "A3", Value: "Widget B"},
                    {Address: "B3", Value: 2500.75},
                    {Address: "B4", Formula: "SUM(B2:B3)"},
                },
                ColWidths: map[string]float64{
                    "A": 20,
                    "B": 15,
                },
            },
        },
    }

    return excelrecreator.QuickRecreate(metadata, "sales_report.xlsx")
}
```

## Performance Considerations

- Large files with many styles may take time to process
- Use `SkipEmptyCells: true` to improve performance
- Consider disabling features you don't need
- Style mapping is cached for efficiency

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

This project uses the [excelize](https://github.com/xuri/excelize) library, which is licensed under the BSD 3-Clause License.

## Dependencies

- [excelize](https://github.com/xuri/excelize) - Excel file manipulation
- [excelmetadata](https://github.com/prongbang/excelmetadata) - Metadata extraction

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/prongbang/excelrecreator).
