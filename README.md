# Excel BOM Formatter

A VBA macro that automatically formats Excel Bill of Materials (BOM) spreadsheets and creates a CAD-ready copy with standardized formatting.

## Features

**Automatic Formatting**: Applies consistent font styles, sizes, and alignment to BOM sections, headers, and data rows
**Multiple Section Triggers**: User can specify multiple texts to trigger section header formatting (e.g., "SECTION", "PANEL SCHEDULE")
**CAD Sheet Generation**: Creates a duplicate sheet with:
  - All text converted to uppercase
  - Optional column removal (e.g., "BPP SKU")
  - No fill colors on cells
  - Text wrapping and auto-fit rows
**Configurable**: Easy-to-customize settings for sheet names, columns, formatting rules, and section header triggers
**Smart Overwriting**: Automatically replaces existing CAD sheets without prompting

## How It Works

The macro processes the active Excel sheet in three steps:

1. **Format Original Sheet**: Applies proper formatting to headers, sections, and data rows
2. **Create CAD Copy**: Duplicates the formatted sheet and applies CAD-specific transformations
3. **Final Cleanup**: Removes the specified column, converts text to uppercase, and clears cell colors

## Installation

### Global Macro Setup (Recommended)

For detailed instructions on setting up this macro as a global macro available in all workbooks, see [Instructions.md](Instructions.md).

**Quick steps:**
1. Press `Alt + F11` to open the VBA Editor
2. Right-click `VBAProject (PERSONAL.XLSB)` → **Import File**
3. Select `FormatBom.bas`
4. Press `Alt + F8`, select the macro, and click **Options** to assign a shortcut key

## Usage

1. Open your Excel BOM file
2. Make sure the sheet you want to format is active
3. Run the macro:
   - Press your assigned shortcut key (e.g., `Ctrl + Q`)
   - Or press `Alt + F8`, select `FormatBom`, and click **Run**

The macro will:
- Format the current sheet with proper fonts and alignment
- Create a new "CAD" sheet with uppercase text and no fill colors
- Optionally rename the original sheet to match the filename

## Configuration

You can customize the macro by editing the **USER CONFIGURATION** section at the top of `FormatBom.bas`:

```vba
Dim sectionTexts As Variant: sectionTexts = Array("SECTION", "SHIPPED")

Dim columnToRemove As String: columnToRemove = "BPP SKU"
Dim activateOriginal As Boolean: activateOriginal = True
Dim newSheetName As String: newSheetName = "CAD"
Dim sheetSuffix As String: sheetSuffix = ""

Dim billOfMaterialText As String: billOfMaterialText = "BILL OF MATERIAL"
Dim tableHeaders As Variant: tableHeaders = Array("ITEM#", "QTY", "BPP SKU", "MFR PART #", "MANUFACTURER", "DESCRIPTION")

Dim centerAlignedColumns As Variant: centerAlignedColumns = Array("ITEM#", "QTY", "BPP SKU")
Dim leftAlignedColumns As Variant: leftAlignedColumns = Array("MFR PART #", "MANUFACTURER", "DESCRIPTION")
```

### Key Settings

| Setting                | Description                                                      | Default                                                      |
|------------------------|------------------------------------------------------------------|--------------------------------------------------------------|
| **`sectionTexts`**         | **Array of texts that trigger section formatting (14pt bold centered)** | `("SECTION", "SHIPPED")` (add more as needed)            |
| `columnToRemove`       | Column header to remove from CAD sheet                          | `"BPP SKU"`                                                  |
| `activateOriginal`     | Return focus to original sheet when done                        | `True`                                                       |
| `newSheetName`         | Name for the CAD-ready sheet                                     | `"CAD"`                                                      |
| `sheetSuffix`          | Suffix to add to original sheet name                            | `""` (uses filename)                                         |
| `billOfMaterialText`   | Text that triggers title formatting (16pt bold centered)         | `"BILL OF MATERIAL"`                                         |
| `tableHeaders`         | Array of expected column headers (11pt bold)                    | `("ITEM#", "QTY", "BPP SKU", ...)`                        |
| `centerAlignedColumns` | Columns to center-align                                          | `("ITEM#", "QTY", "BPP SKU")`                              |
| `leftAlignedColumns`   | Columns to left-align                                            | `("MFR PART #", "MANUFACTURER", "DESCRIPTION")`             |