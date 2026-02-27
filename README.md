# Sheety

Sheety is an Excel-like spreadsheet application for Crystal that compiles spreadsheet definitions (YAML or Excel) into standalone, interactive TUI binaries.

> **⚠️ Note:** This is experimental software. We're compiling spreadsheets into standalone binaries. It's ridiculous, it's probably unnecessary, and it absolutely works.

**🌐 Check out the website:** [sheety.ralsina.me](https://sheety.ralsina.me)

## Features

- **Formula Parser**: Parses a comprehensive subset of Excel formulas
  - Arithmetic: `+`, `-`, `*`, `/`, `^`
  - Comparison: `=`, `<`, `>`, `<=`, `>=`, `<>`
  - Text: `&` (concatenation)
  - Logical: `IF`, `AND`, `OR`, `NOT`, `IFS`, `SWITCH`
  - **Math & Statistical**: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `MEDIAN`, `STDEV`, `STDEV.P`, `VAR.S`, `VAR.P`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `INT`, `ABS`, `POWER`, `SQRT`, `MOD`, `RAND`, `RANDBETWEEN`
  - **Text Functions**: `CONCAT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `PROPER`, `FIND`, `SEARCH`, `SUBSTITUTE`, `TEXT`, `VALUE`, `CLEAN`, `EXACT`, `REPT`
  - **Date & Time**: `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `DATEDIF`, `EOMONTH`
  - **Conditional**: `COUNTIF`, `SUMIF`
  - **Lookup**: `VLOOKUP`, `HLOOKUP`, `INDEX`
  - Cell references: `A1`, `Sheet2!B5`
  - Ranges: `A1:B5`, `Sheet1!A1:A10`

- **Format Support**: Works with multiple spreadsheet formats
  - **YAML** - Human-readable text format for version control
  - **Excel (.xlsx)** - Import and export Excel files
  - Convert between any formats

- **Hash-Based Build Caching**: Content-based caching makes re-running unchanged spreadsheets ~400x faster
  - First run: ~3 seconds (compilation)
  - Cached runs: ~0.007 seconds (instant startup)
  - Automatically recompiles when source content changes

- **TUI Interface**: Modern terminal UI with spreadsheet-style grid
  - 1000x1000 grid with column headers (A-Z) and row numbers
  - Active cell highlighting
  - Formula bar showing current cell's formula or value
  - Multiple sheet support with Tab switching

- **Interactive Editing**: Full cell editing with live recalculation
  - Edit value cells by typing
  - Edit formulas with real-time cursor
  - Auto-recalculation of dependent formulas
  - Formula editing triggers automatic binary rebuild

- **Mouse Support**: Click to select, double-click to edit, wheel scrolling

- **State Persistence**: Saves and restores UI state
  - Remembers cursor position across sessions
  - Saves all changes back to source file
  - UUID-based state management survives rebuilds

- **In-Process Rebuilding**: Formula edits rebuild the binary from within the TUI

## Installation

```bash
git clone https://github.com/ralsina/sheety.git
cd sheety
shards install
shards build
```

## Usage

### Command Line

```bash
# Compile and run a spreadsheet (auto-creates if doesn't exist)
./bin/sheety my_sheet.yaml

# Convert between formats
./bin/sheety data.yaml --save-to=data.xlsx     # YAML to Excel
./bin/sheety data.xlsx --save-to=data.yaml     # Excel to YAML
./bin/sheety data.yaml --save-to=data.cr       # Generate Crystal source
./bin/sheety data.yaml --save-to=data.sheety   # Compile to binary

# Works with Excel files directly
./bin/sheety workbook.xlsx
```

### Spreadsheet YAML Format

```yaml
Sheet1:
  A1:
    value: 100
  A2:
    value: 200
  A3:
    formula: "=SUM(A1:A2)"
  B1:
    value: "Hello"
  B2:
    formula: "=CONCAT(B1, \" World\")"

Sheet2:
  A1:
    formula: "=Sheet1!A3*2"
```

### TUI Controls

**Navigation:**
- Arrow keys, Home/End, Page Up/Down
- Tab to switch sheets
- Click to select, double-click to edit

**Editing:**
- Enter to edit current cell
- S to save spreadsheet
- Q to quit

## Development

```bash
# Run tests
crystal spec

# Build
shards build

# Format conversion
./bin/sheety examples/test.yaml --save-to=output.xlsx
```

## Current Status

- ✅ Excel formula parsing
- ✅ YAML and Excel format support with full roundtrip conversion
- ✅ Code generation and compilation to standalone binaries
- ✅ TUI interface with interactive editing
- ✅ In-process rebuilding on formula changes
- ✅ State persistence across rebuilds
- ✅ Auto-creation of new spreadsheets
- ✅ Hash-based build caching

## License

MIT

## Contributors

- [Roberto Alsina](https://github.com/ralsina) - creator and maintainer
