# Sheety - Excel Formula Parser (Crystal Port)

A Crystal port of the Python `formulas` library for parsing Excel formulas.

## Features Implemented

### Core Functionality
- ✅ Formula parsing with tokenization
- ✅ Abstract Syntax Tree (AST) generation
- ✅ Operator precedence handling (shunting-yard algorithm)
- ✅ Parentheses for grouping expressions
- ✅ Comprehensive error handling

### Supported Tokens

#### Literals
- **Numbers**: Integers, decimals, scientific notation (`1`, `3.14`, `1E+10`)
- **Strings**: Double-quoted strings with escape support (`"hello"`, `"a""b"`)
- **Booleans**: `TRUE`, `FALSE` (case-insensitive)
- **Errors**: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NUM!`, `#NAME?`, `#N/A`

#### References
- **Cell references**: `A1`, `$A1`, `A$1`, `$A$1`
- **Ranges**: `A1:B5`, `A:B`, `1:10`
- **Sheet references**: `Sheet1!A1`, `'My Sheet'!A1`, `Sheet1!A1:B5`
- **Named ranges**: `MyRange`, `SalesData`, `Total_Sales`

#### Operators
- **Arithmetic**: `+`, `-`, `*`, `/`, `^` (exponentiation)
- **Comparison**: `=`, `<`, `>`, `<=`, `>=`, `<>`
- **Text concatenation**: `&`
- **Unary**: `+`, `-`, `%`
- **Intersection**: ` ` (space)

#### Functions
- **Function calls**: `SUM()`, `AVERAGE()`, `IF()`, etc.
- **Nested functions**: `SUM(A1, MAX(B1:B5))`
- **Multiple arguments**: `IF(A1>0, 1, 0)`
- **No arguments**: `PI()`

#### Array Constants
- **Simple arrays**: `{1,2,3}`
- **2D arrays**: `{{1,2},{3,4}}`
- **Mixed types**: `{1, "text", TRUE}`
- **Nested arrays**: Supported

## Test Coverage

**186 examples**, all passing:
- 71 original Sheety tests
- 56 parser tests adapted from Python formulas
- 59 token tests adapted from Python formulas

## Project Structure

```
src/sheety/
├── ast.cr              # AST node classes
├── ast_builder.cr       # AST construction
├── parser.cr            # Main formula parser
├── token.cr             # Base token class
├── errors.cr            # Error classes
└── tokens/
    ├── operand.cr       # Number, String, Boolean, Error, Range, Array
    ├── operator.cr       # All operator types
    ├── parenthesis.cr    # Parenthesis handling
    ├── function_call.cr  # Function name tokens
    ├── argument_separator.cr  # Function argument commas
    ├── array_constant.cr # Array constant tokens
    └── named_range.cr    # Named range tokens

spec/
├── spec_helper.cr
├── sheety_spec.cr              # Core functionality tests
├── cell_refs_spec.cr           # Cell reference tests
├── functions_spec.cr            # Function call tests
├── intersection_spec.cr         # Intersection operator tests
├── named_ranges_spec.cr         # Named range tests
├── union_spec.cr                # Union operator tests
├── array_constants_spec.cr      # Array constant tests
├── formulas_parser_spec.cr      # Adapted from Python test_parser.py
└── formulas_token_spec.cr       # Adapted from Python test_tokens.py
```

## Usage

```crystal
require "sheety"

# Parse a formula
ast = Sheety.parse_to_ast("=SUM(A1:B5) * 2")

# The AST contains the formula structure
puts ast.expr  # => "(SUM(A1:B5) * 2)"

# Check the type
case ast
when Sheety::AST::FunctionCall
  puts "Function: #{ast.function_name}"
when Sheety::AST::BinaryOp
  puts "Operator: #{ast.operator}"
when Sheety::AST::Number
  puts "Number: #{ast.value}"
end
```

## Missing Features (from Python formulas)

- R1C1 reference style
- `@` implicit intersection operator
- LAMBDA/LET functions
- Array row separator (`;`) for 2D arrays
- #GETTING_DATA error
- External workbook references
- Sheet-prefixed named ranges

## Building

```bash
shards build
crystal spec
```

## Status

✅ Fully functional parser with comprehensive test coverage
🎯 186 tests passing
📚 Adapted from Python formulas library test suite
