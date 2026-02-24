require "./sheety/errors"
require "./sheety/token"
require "./sheety/ast"
require "./sheety/ast_builder"
require "./sheety/parser"
require "./sheety/tokens/operand"
require "./sheety/tokens/operator"
require "./sheety/tokens/parenthesis"
require "./sheety/functions/registry"
require "./sheety/evaluator"
require "./sheety/code_generator"
require "./sheety/dependency_extractor"
require "./sheety/croupier_generator"
require "./sheety/standalone_generator"
require "./sheety/cli"

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

  # Generate Crystal code from an Excel formula AST
  #
  # Example:
  # ```
  # ast = Sheety.parse_to_ast("=SUM(A1:A5)")
  # code = Sheety.generate_code(ast)
  # puts code # => Something like "Sheety::Functions.sum([...])"
  # ```
  def self.generate_code(node : AST::Node, sheet : String? = nil) : String
    generator = CodeGenerator.new
    generator.generate(node, CodeGenerator::Context.new(sheet))
  end

  # Generate a complete proc body for a formula
  #
  # Example:
  # ```
  # proc_body = Sheety.generate_proc_body("=SUM(A1:A5)")
  # ```
  def self.generate_proc_body(formula : String, sheet : String? = nil) : String
    generator = CodeGenerator.new
    generator.generate_proc_body(formula, CodeGenerator::Context.new(sheet))
  end

  # Extract cell dependencies from a formula
  #
  # Example:
  # ```
  # deps = Sheety.extract_dependencies("=SUM(A1:B5)+C1")
  # puts deps.to_a # => ["A1", "A2", ..., "B5", "C1"]
  # ```
  def self.extract_dependencies(formula : String, sheet : String? = nil) : Set(String)
    extractor = DependencyExtractor.new
    extractor.extract_from_formula(formula, sheet)
  end

  # Create a new CroupierGenerator for generating Croupier tasks
  #
  # Example:
  # ```
  # gen = Sheety.croupier_generator
  # gen.add_formula("C1", "=SUM(A1:A5)")
  # gen.add_formula("D1", "=C1*2")
  # gen.register_tasks
  # ```
  def self.croupier_generator : CroupierGenerator
    CroupierGenerator.new
  end
end

# CLI entry point - run when executed directly
Sheety::CLI.run(ARGV) if PROGRAM_NAME.includes?("sheety")
