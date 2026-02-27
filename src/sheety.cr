require "./sheety/errors"
require "./sheety/token"
require "./sheety/ast"
require "./sheety/ast_builder"
require "./sheety/parser"
require "./sheety/tokens/operand"
require "./sheety/tokens/operator"
require "./sheety/tokens/parenthesis"
require "./sheety/functions/registry"
require "./sheety/code_generator"
require "./sheety/dependency_extractor"
require "./sheety/croupier_generator"
require "./sheety/embedded_files"
require "./sheety/cli"

# TODO: Write documentation for `Sheety`
module Sheety
  VERSION = "0.1.0"

  # Parse an Excel formula and return the AST root node
  def self.parse_to_ast(formula : String) : AST::Node
    parser = Parser.new
    _, builder = parser.ast(formula)
    builder.root
  end
end

# CLI entry point - run when executed directly
Sheety::CLI.run(ARGV) if PROGRAM_NAME.includes?("sheety")
