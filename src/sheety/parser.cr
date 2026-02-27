require "big"
require "./errors"
require "./ast_builder"
require "./tokens/operand"
require "./tokens/operator"
require "./tokens/parenthesis"
require "./tokens/function_call"
require "./tokens/argument_separator"
require "./tokens/array_constant"
require "./tokens/named_range"

module Sheety
  # Main parser class for Excel formulas
  class Parser
    # Regex to check if a string is a valid formula
    FORMULA_CHECK = /^\s*=\s*(?P<name>.+)|^\s*{\s*=\s*(?P<name>.+)\s*}/i

    property context : Hash(String, BigFloat | String)?

    def initialize(@context : Hash(String, BigFloat | String)? = nil)
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
        Tokens::ArrayConstant, # Must come before Function for {...}
        Tokens::StringToken,
        Tokens::Boolean,
        Tokens::Function,   # Must come before other tokens that start with letters
        Tokens::NamedRange, # Must come before Range to distinguish from cell refs
        Tokens::Range,      # Must come before Number to match ranges like 1:10
        Tokens::Number,
        Tokens::ArgumentSeparator, # For function argument commas
        Tokens::ComparisonOperator,
        Tokens::ConcatOperator, # Text concatenation (&)
        Tokens::ArithmeticOperator,
        Tokens::PercentOperator,
        Tokens::ColonOperator,
        # Tokens::SeparatorOperator,  # TODO: Re-enable for array formulas/union
        Tokens::IntersectOperator,
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
      elsif filter_class == Tokens::ArrayConstant
        if Tokens::ArrayConstant.match?(expr)
          Tokens::ArrayConstant.new(expr, @context)
        end
      elsif filter_class == Tokens::NamedRange
        if Tokens::NamedRange.match?(expr)
          Tokens::NamedRange.new(expr, @context)
        end
      elsif filter_class == Tokens::Function
        if Tokens::Function.match?(expr)
          Tokens::Function.new(expr, @context)
        end
      elsif filter_class == Tokens::ArgumentSeparator
        if Tokens::ArgumentSeparator.match?(expr)
          Tokens::ArgumentSeparator.new(expr, @context)
        end
      elsif filter_class == Tokens::Number
        if Tokens::Number.match?(expr)
          Tokens::Number.new(expr, @context)
        end
      elsif filter_class == Tokens::Range
        if Tokens::Range.match?(expr)
          Tokens::Range.new(expr, @context)
        end
      elsif filter_class == Tokens::ComparisonOperator
        if Tokens::ComparisonOperator.match?(expr)
          Tokens::ComparisonOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::ArithmeticOperator
        if Tokens::ArithmeticOperator.match?(expr)
          Tokens::ArithmeticOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::ConcatOperator
        if Tokens::ConcatOperator.match?(expr)
          Tokens::ConcatOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::PercentOperator
        if Tokens::PercentOperator.match?(expr)
          Tokens::PercentOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::ColonOperator
        if Tokens::ColonOperator.match?(expr)
          Tokens::ColonOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::SeparatorOperator
        if Tokens::SeparatorOperator.match?(expr)
          Tokens::SeparatorOperator.new(expr, @context)
        end
      elsif filter_class == Tokens::IntersectOperator
        if Tokens::IntersectOperator.match?(expr)
          Tokens::IntersectOperator.new(expr, @context)
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
