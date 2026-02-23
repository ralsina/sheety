require "./sheety/errors"
require "./sheety/token"
require "./sheety/ast_builder"
require "./sheety/parser"
require "./sheety/tokens/operand"
require "./sheety/tokens/operator"
require "./sheety/tokens/parenthesis"

# TODO: Write documentation for `Sheety`
module Sheety
  VERSION = "0.1.0"

  # Parse an Excel formula and return an executable function
  #
  # Example:
  # ```
  # func = Sheety.parse("=1 + 2 * 3")
  # result = func.call({})  # => 7.0
  # ```
  def self.parse(formula : String) : (Hash(String, Float64 | String) -> Float64 | String)
    parser = Parser.new
    parser.parse(formula)
  end

  # Parse an Excel formula and evaluate it with given inputs
  #
  # Example:
  # ```
  # result = Sheety.evaluate("=A1 + B1", {"A1" => 10.0, "B1" => 5.0}) # => 15.0
  # ```
  def self.evaluate(formula : String, inputs : Hash(String, Float64 | String) = Hash(String, Float64 | String).new) : Float64 | String
    func = parse(formula)
    func.call(inputs)
  end
end
