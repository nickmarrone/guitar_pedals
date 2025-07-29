# Aion FX BOM Merge Utility

This utility script processes and merges Bill of Materials (BOM) files from Aion FX guitar pedal projects into a consolidated Excel workbook.

## Overview

The `aion_fx_bom_merge.py` script takes multiple Aion FX BOM Excel files and combines them into a single Excel workbook with the following features:

- **Individual Project Sheets**: Each input BOM becomes a separate sheet in the output file
- **Combined Summary Sheet**: Aggregates all parts across all projects with total counts
- **Smart Categorization**: Automatically categorizes components by type (Resistors, Capacitors, Diodes, etc.)
- **Intelligent Sorting**: Sorts components by type and value for easy reading
- **Inventory Integration**: Optional inventory checking with color-coded highlighting for missing/low-stock parts

## Features

### Component Categorization
The script automatically categorizes components into these types:
- Resistors
- Capacitors  
- Diodes (including LEDs)
- Transistors
- ICs (Integrated Circuits)
- Potentiometers
- Switches
- Connectors
- Other

### Smart Value Parsing
- Handles Euro-style notation (e.g., "2K2" → "2.2K")
- Converts capacitor values to consistent format
- Sorts resistors by resistance value
- Sorts capacitors by capacitance value

### Inventory Integration
When an inventory file is provided, the script:
- Checks resistor availability against "TH Resistors" sheet
- Checks capacitor availability against "TH Capacitors" sheet (ceramic, film, electrolytic)
- Highlights missing parts in **pink**
- Highlights low-stock parts in **orange**

### Data Cleaning
- Removes excluded components (IC sockets, enclosures, dust covers)
- Filters out rows with missing essential data
- Normalizes part numbers and descriptions
- Groups duplicate components across projects

## Usage

### Basic Usage
```bash
python aion_fx_bom_merge.py --in "project1.xlsx" "project2.xlsx" --out "combined_bom.xlsx"
```

### With Inventory Checking
```bash
python aion_fx_bom_merge.py --in "project1.xlsx" "project2.xlsx" --out "combined_bom.xlsx" --inventory "inventory.xlsx"
```

### Command Line Arguments
- `--in`: Input BOM Excel files (one or more)
- `--out`: Output Excel filename
- `--inventory`: Optional inventory Excel file for stock checking

## Input File Requirements

### BOM Files
- Excel (.xlsx) format
- Should be the "Mouser parts spreadsheet" linked from each Aion FX project page
- Should contain columns: Part, Value, Description, Notes
- Multiple sheets per file are supported (except "Instructions" and "Combined" sheets)

### Inventory File (Optional)
- Excel (.xlsx) format
- **TH Resistors** sheet: Column A = value, Column B = status
- **TH Capacitors** sheet: 
  - Columns A-B: Ceramic capacitors
  - Columns D-E: Film capacitors  
  - Columns F-G: Electrolytic capacitors

#### Example Inventory File Structure

**TH Resistors Sheet:**
| Value | Status |
|-------|--------|
| 10r   |        |
| 100r  | few    |
| 1k    |        |
| 10k   |        |
| 100k  |        |
| 1m    |        |

**TH Capacitors Sheet:**
| Ceramic Value | Status | | Film Value | Status | | Electrolytic Value | Status |
|---------------|--------|-|------------|--------|-|-------------------|--------|
| 10p           |        | | 1n         |        | | 1u               |        |
| 22p           | few    | | 1n5        |        | | 10u              |        |
| 100p          |        | | 2n2        |        | | 47u              | few    |
| 1n            |        | | 10n        |        | | 100u             |        |
| 10n           |        | | 100n       |        | | 220u             |        |
| 100n          | few    | | 1u         |        | | 470u             |        |

**Status Values:**
- Empty: Plenty in stock
- `few`: Low stock (will be highlighted in orange)
- Missing entry: Not in stock (will be highlighted in pink)

## Output Structure

The generated Excel file contains:
1. **Individual project sheets**: One per input BOM file
2. **Combined sheet**: Aggregated view of all components across all projects
3. **Auto-formatted columns**: Column widths adjusted to content
4. **Color-coded highlighting**: Missing/low-stock parts highlighted (if inventory provided)

## Dependencies

- `pandas`: Data manipulation and Excel I/O
- `openpyxl`: Excel file handling
- `natsort`: Natural sorting of part numbers
- `argparse`: Command line argument parsing

## Example Output

The script creates a professional BOM workbook suitable for:
- Component purchasing
- Inventory management
- Project planning
- Cost analysis

Each sheet is properly formatted with auto-sized columns and clear component organization. 