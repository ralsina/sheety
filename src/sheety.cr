require "./sheety/errors"
require "./sheety/token"
require "./sheety/ast"
require "./sheety/ast_builder"
require "./sheety/parser"
require "./sheety/tokens/operand"
require "./sheety/tokens/operator"
require "./sheety/tokens/parenthesis"

# TODO: Write documentation for `Sheety`
module Sheety
  VERSION = "0.1.0"

  # Parse an Excel formula and return the AST root node
  #
  # Example:
  # ```
  # ast = Sheety.parse_to_ast("=1 + 2 * 3")
  # puts ast.expr # => "(1 + (2 * 3))"
  # ```
  def self.parse_to_ast(formula : String) : AST::Node
    parser = Parser.new
    _, builder = parser.ast(formula)
    builder.root
  end

  # Parse an Excel formula and return both tokens and AST
  #
  # Example:
  # ```
  # tokens, builder = Sheety.parse("=1 + 2 * 3")
  # puts builder.root.expr # => "(1 + (2 * 3))"
  # ```
  def self.parse(formula : String) : Tuple(Array(Token), AstBuilder)
    parser = Parser.new
    parser.ast(formula)
  end
end
