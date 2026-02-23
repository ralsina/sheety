require "./errors"
require "./ast_builder"
require "./tokens/operand"
require "./tokens/operator"
require "./tokens/parenthesis"

module Sheety
  # Main parser class for Excel formulas
  class Parser
    # Regex to check if a string is a valid formula
    FORMULA_CHECK = /^\s*=\s*(?P<name>.+)|^\s*{\s*=\s*(?P<name>.+)\s*}/i

    property context : Hash(String, Float64 | String)?

    def initialize(@context : Hash(String, Float64 | String)? = nil)
    end

    # Check if a value is a formula
    def formula?(value : String) : Bool
      FORMULA_CHECK.match(value) != nil
    end

    # Parse a formula string and return (tokens, ast_builder)
    def ast(expression : String) : Tuple(Array(Token), AstBuilder)
      # Normalize expression
      expr = expression.gsub("\n", "")

      # Check if it's a formula
      match = FORMULA_CHECK.match(expr)
      unless match
        raise FormulaError.new("Not a formula: #{expression}")
      end

      expr_body = match["name"]?
      unless expr_body
        raise FormulaError.new("Could not extract formula body from: #{expression}")
      end

      # Token filters in order of priority
      filters = [
        Tokens::ErrorToken,
        Tokens::StringToken,
        Tokens::Boolean,
        Tokens::Number,
        Tokens::ComparisonOperator,
        Tokens::ArithmeticOperator,
        Tokens::PercentOperator,
        Tokens::Parenthesis,
      ]

      tokens = Array(Token).new
      stack = Array(Token).new
      builder = AstBuilder.new

      # Add implicit opening parenthesis
      Tokens::Parenthesis.new("(").ast(tokens, stack, builder)

      # Parse tokens
      while expr_body.size > 0
        matched = false

        filters.each do |filter_class|
          begin
            if token = try_match_token(filter_class, expr_body)
              token.ast(tokens, stack, builder)
              expr_body = expr_body[token.end_match..-1]
              matched = true
              break
            end
          rescue TokenError
            # Try next filter
          rescue ex : FormulaError
            raise FormulaError.new("Error parsing: #{expression}")
          end
        end

        unless matched
          raise FormulaError.new("Could not parse: #{expr_body}")
        end
      end

      # Add implicit closing parenthesis
      Tokens::Parenthesis.new(")").ast(tokens, stack, builder)

      # Remove implicit parentheses
      tokens = tokens[1..-2]? || Array(Token).new

      # Pop remaining operators from stack
      while stack.size > 0
        token = stack.pop?
        if token.is_a?(Tokens::Parenthesis)
          raise ParenthesesError.new
        end
        if token
          builder.append(token)
        end
      end

      # Finish building
      builder.finish

      {tokens, builder}
    end

    private def try_match_token(filter_class : Token.class, expr : String) : Token?
      # Use if/elsif instead of case for class comparison
      if filter_class == Tokens::ErrorToken
        if Tokens::ErrorToken.match?(expr)
          Tokens::ErrorToken.new(expr, @context)
        end
      elsif filter_class == Tokens::StringToken
        if Tokens::StringToken.match?(expr)
          Tokens::StringToken.new(expr, @context)
        end
      elsif filter_class == Tokens::Boolean
        if Tokens::Boolean.match?(expr)
          Tokens::Boolean.new(expr, @context)
        end
      elsif filter_class == Tokens::Number
        if Tokens::Number.match?(expr)
          Tokens::Number.new(expr, @context)
        end
      elsif filter_class == Tokens::ComparisonOperator
        if Tokens::ComparisonOperator.match?(expr)
          Tokens::ComparisonOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::ArithmeticOperator
        if Tokens::ArithmeticOperator.match?(expr)
          Tokens::ArithmeticOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::PercentOperator
        if Tokens::PercentOperator.match?(expr)
          Tokens::PercentOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::Parenthesis
        if Tokens::Parenthesis.match?(expr)
          Tokens::Parenthesis.new(expr, @context)
        end
      else
        nil
      end
    end
  end
end
