# Sheety

Sheety is an Excel-like spreadsheet library for Crystal that parses formulas, evaluates expressions, and generates standalone binaries using Croupier for task-based dependency tracking.

## Features

- **Excel Formula Parser**: Parses a comprehensive subset of Excel formulas
  - Arithmetic: `+`, `-`, `*`, `/`, `^`
  - Comparison: `=`, `<`, `>`, `<=`, `>=`, `<>`
  - Text: `&` (concatenation)
  - Logical: `IF`, `AND`, `OR`, `NOT`, `IFS`, `SWITCH`
  - **Math & Statistical**:
    - Basic: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `MEDIAN`
    - Statistics: `STDEV`, `STDEV.P`, `VAR.S`, `VAR.P`
    - Rounding: `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `INT`, `ABS`
    - Advanced: `POWER`, `SQRT`, `MOD`, `RAND`, `RANDBETWEEN`
  - **Text Functions**: `CONCAT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `PROPER`, `FIND`, `SEARCH`, `SUBSTITUTE`, `TEXT`, `VALUE`, `CLEAN`, `EXACT`, `REPT`
  - **Date & Time**: `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `DATEDIF`, `EOMONTH`
  - **Conditional**: `COUNTIF`, `SUMIF`
  - **Lookup**: `VLOOKUP`, `HLOOKUP`, `INDEX`
  - Cell references: `A1`, `Sheet2!B5`
  - Ranges: `A1:B5`, `Sheet1!A1:A10`

- **Dependency Tracking**: Uses Croupier to track dependencies and only recalculate affected cells

- **TUI Interface**: Modern terminal UI with Lotus 1-2-3 style grid interface
  - 1000x1000 grid display with headers and row numbers
  - Column headers (A, B, C...) and row numbers (1, 2, 3...)
  - Active cell highlighting
  - Formula bar showing current cell's formula or value
  - Status bar with navigation hints
  - Multiple sheet support with Tab switching

- **Interactive Editing**: Full cell editing capabilities
  - Edit value cells by typing new values
  - Edit formula cells directly with real-time cursor
  - Auto-recalculation of dependent formulas on any change
  - Formula editing triggers automatic recompile and restart
  - Visual cursor positioning in edit mode

- **Mouse Support**: Intuitive mouse navigation
  - Click to select cells
  - Double-click to edit cells
  - Mouse wheel scrolling through rows

- **State Persistence**: Saves and restores UI state
  - Remembers cursor position (sheet and cell) across sessions
  - Saves modified values and formulas back to YAML
  - Scans entire grid to capture all changes

## Installation

Add the dependency to your `shard.yml`:

```yaml
dependencies:
  sheety:
    github: ralsina/sheety
```

Then run:

```bash
shards install
```

## Usage

### Command Line

Sheety provides a CLI for compiling spreadsheet definitions into standalone binaries:

```bash
# Compile a spreadsheet YAML to a Crystal binary (auto-builds and runs)
sheety compile examples/test_sheet.yaml

# The binary is automatically built and launched with the TUI
# Press Q to exit
```

### TUI Controls

**Navigation:**
- **Arrow keys** - Move cursor up/down/left/right
- **Home/End** - Jump to first/last column or row
- **Page Up/Down** - Scroll by page
- **Tab** - Switch between sheets
- **Mouse wheel** - Scroll up/down
- **Click** - Select cell
- **Double-click** - Edit cell

**Editing:**
- **Enter** - Edit current cell
- **Escape** - Cancel edit / Exit TUI
- **S** - Save spreadsheet to YAML
- **Q** - Quit

**Edit Mode:**
- **Arrow keys** - Move cursor within text
- **Home/End** - Jump to start/end of text
- **Backspace** - Delete character before cursor
- **Delete** - Delete character at cursor
- **Enter** - Save changes
- **Escape** - Cancel edit

### Spreadsheet YAML Format

Define your spreadsheet in YAML:

```yaml
Sheet1:
  A1:
    value: 100
  A2:
    value: 200
  A3:
    value: 300
  A4:
    formula: "=SUM(A1:A3)"
  A5:
    formula: "=AVERAGE(A1:A3)"
  B1:
    value: "Hello"
  B2:
    value: "World"
  B3:
    formula: "=CONCAT(B1, \" \", B2)"
  C1:
    value: 5
  C2:
    value: 3
  C3:
    formula: "=IF(C1>C2, \"Yes\", \"No\")"

Sheet2:
  A1:
    value: 100
  A2:
    formula: "=Sheet1!A4*2"
```

### Programmatic Usage

```crystal
require "sheety"

generator = Sheety::CroupierGenerator.new

# Add formulas
generator.add_formula("A1", "=SUM(B1:B10)")
generator.add_formula("C5", "=AVERAGE(A1:A100)", "Sheet2")

# Generate source code with initial values
initial_values = {
  "Sheet1!B1" => 10.0,
  "Sheet1!B2" => 20.0,
}
source = generator.generate_source(initial_values, interactive: true)

# Write and compile
File.write("my_sheet.cr", source)
`crystal build my_sheet.cr`
```

## Development

### Running Tests

```bash
crystal spec
```

### Building

```bash
shards build
```

### Examples

See the `examples/` directory for sample spreadsheet definitions.

## Current Status

- ✅ Excel formula parsing with 278+ passing tests
- ✅ AST generation for all formula types
- ✅ Code generation for Croupier tasks
- ✅ Cell reference and range support
- ✅ Multi-sheet support
- ✅ 50+ Excel function implementations (math, statistical, text, date, conditional, lookup)
- ✅ TUI interface with Termisu
- ✅ Interactive cell editing with formula and value support
- ✅ Mouse support (selection, editing, scrolling)
- ✅ State persistence (cursor position, values, formulas)
- ✅ Automatic recompilation on formula changes
- ✅ Dependency tracking with Croupier

## License

MIT

## Contributors

- [Roberto Alsina](https://github.com/ralsina) - creator and maintainer
