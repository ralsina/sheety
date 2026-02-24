# Sheety

Sheety is an Excel-like spreadsheet library for Crystal that parses formulas, evaluates expressions, and generates standalone binaries using Croupier for task-based dependency tracking.

## Features

- **Excel Formula Parser**: Parses a comprehensive subset of Excel formulas
  - Arithmetic: `+`, `-`, `*`, `/`, `^`
  - Comparison: `=`, `<`, `>`, `<=`, `>=`, `<>`
  - Text: `&` (concatenation)
  - Logical: `IF`, `AND`, `OR`, `NOT`
  - Functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `ROUND`, `ABS`, `CONCAT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`
  - Cell references: `A1`, `Sheet2!B5`
  - Ranges: `A1:B5`, `Sheet1!A1:A10`

- **Dependency Tracking**: Uses Croupier to track dependencies and only recalculate affected cells

- **Interactive Mode**: Generated binaries support interactive REPL for modifying cell values

- **Spreadheet Display**: Shows results in a proper grid layout using Tablo

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
# Compile a spreadsheet YAML to a Crystal binary
sheety compile examples/test_sheet.yaml

# Run the generated binary (interactive mode by default)
./test_sheet
```

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

### Interactive Mode

When you run the generated binary, you enter interactive REPL mode:

```
=== Spreadsheet Results ===
[... table display ...]

=== Interactive Mode ===
Enter cell assignments (e.g., A1=123, Sheet2!B5=hello)
Commands: 'quit' or 'exit' to quit, 'show' to refresh display

> A1=999
Set Sheet1!A1 = 999
[... updated table with recalculated formulas ...]

> quit
Goodbye!
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

- ✅ Excel formula parsing with 186 passing tests
- ✅ AST generation for all formula types
- ✅ Code generation for Croupier tasks
- ✅ Cell reference and range support
- ✅ Multi-sheet support
- ✅ Excel function implementations
- ✅ Interactive REPL mode
- ✅ Spreadsheet grid display with Tablo
- ✅ Dependency tracking with Croupier
- ✅ Automatic recalculation on value changes

## License

MIT

## Contributors

- [Roberto Alsina](https://github.com/ralsina) - creator and maintainer
