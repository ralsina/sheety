require "./token"
require "./errors"
require "./tokens/operand"
require "./tokens/operator"

module Sheety
  # Builds an Abstract Syntax Tree (AST) from tokens and compiles to executable code
  class AstBuilder
    property tokens : Array(Token)

    def initialize
      @tokens = Array(Token).new
    end

    def append(token : Token) : Nil
      @tokens << token
    end

    def size : Int32
      @tokens.size
    end

    def [](index : Int32) : Token
      @tokens[index]
    end

    def pop : Token?
      @tokens.pop?
    end

    def finish : Nil
      # Validate that we have exactly one root expression
      if @tokens.size != 1
        raise FormulaError.new("Expected single root expression, got #{@tokens.size} tokens")
      end
    end

    # Compile the AST to an executable function
    def compile(context : Hash(String, Float64 | String)? = nil) : Proc(Hash(String, Float64 | String), Float64 | String)
      root = @tokens.last?
      raise FormulaError.new("No tokens to compile") if root.nil?

      # Return a lambda that evaluates the expression
      ->(inputs : Hash(String, Float64 | String)) {
        evaluate_simple(root, inputs).as(Float64 | String)
      }
    end

    # Simple evaluation for basic arithmetic (no proper AST yet)
    # For the prototype, just evaluate single numeric operands
    private def evaluate_simple(token : Token, inputs : Hash(String, Float64 | String)) : Float64 | String
      case token
      when Tokens::Number
        token.compile.as(Float64)
      when Tokens::StringToken
        token.compile.as(String)
      when Tokens::ErrorToken
        # For errors, we could raise or return a string representation
        token.compile.to_s
      else
        # For operators and more complex expressions, we'd need full AST traversal
        # For now, this is a limitation of the prototype
        raise FormulaError.new("Complex expressions not yet supported: #{token.class}")
      end
    end
  end
end
