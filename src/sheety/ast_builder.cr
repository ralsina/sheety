require "./token"
require "./errors"
require "./ast"
require "./tokens/operand"
require "./tokens/operator"
require "./tokens/parenthesis"

module Sheety
  # Builds an Abstract Syntax Tree (AST) from tokens
  class AstBuilder
    # Stack of AST nodes being built
    property node_stack : Array(AST::Node)

    def initialize
      @node_stack = Array(AST::Node).new
    end

    # Append a token and build AST node
    def append(token : Token) : Nil
      case token
      when Tokens::Operand
        # Operands become leaf nodes
        node = operand_to_ast_node(token)
        @node_stack << node
      when Tokens::Operator
        # Operators pop their operands and create a new node
        n_args = token.n_args
        operands = Array(AST::Node).new

        n_args.times do
          if @node_stack.empty?
            raise FormulaError.new("Not enough operands for operator: #{token.name}")
          end
          operands.unshift(@node_stack.pop)
        end

        # Create the appropriate AST node
        if n_args == 1
          node = AST::UnaryOp.new(token.name, operands[0])
        else
          node = AST::BinaryOp.new(token.name, operands[0], operands[1])
        end

        @node_stack << node
      end
    end

    # Append an AST node directly
    def append(node : AST::Node) : Nil
      @node_stack << node
    end

    def pop : AST::Node?
      @node_stack.pop?
    end

    def finish : Nil
      # Validate that we have exactly one root expression
      if @node_stack.size != 1
        raise FormulaError.new("Expected single root expression, got #{@node_stack.size} nodes")
      end
    end

    # Get the root AST node
    def root : AST::Node
      if @node_stack.size != 1
        raise FormulaError.new("AST does not have exactly one root node")
      end
      @node_stack.last
    end

    private def operand_to_ast_node(token : Tokens::Operand) : AST::Node
      case token
      when Tokens::Number
        value = token.compile.as(BigFloat)
        AST::Number.new(value)
      when Tokens::Boolean
        value = token.bool_value
        AST::Boolean.new(value)
      when Tokens::StringToken
        value = token.compile.as(String)
        AST::StringLiteral.new(value)
      when Tokens::ErrorToken
        error = token.compile
        AST::ErrorValue.new(error.value)
      when Tokens::Range
        sheet = token.sheet_name
        if token.range?
          AST::RangeRef.new(token.name, sheet)
        else
          AST::CellRef.new(token.name, sheet)
        end
      when Tokens::NamedRange
        AST::NamedRef.new(token.name)
      when Tokens::ArrayConstant
        # Parse array elements from the array content
        element_strs = token.parse_elements
        elements = Array(AST::Node).new

        element_strs.each do |elem_str|
          # Try to parse each element as a sub-expression
          elem_str = elem_str.strip

          # Skip empty elements
          if elem_str.empty?
            next
          end

          # Parse the element as a mini-formula
          begin
            elem_parser = Parser.new
            _, elem_builder = elem_parser.ast("=#{elem_str}")
            elem_ast = elem_builder.root
            elements << elem_ast
          rescue ex : Sheety::FormulaError
            # If parsing fails, create a placeholder
            # Check if it's a number
            if elem_str =~ /^-?\d+(\.\d+)?$/
              elements << AST::Number.new(BigFloat.new(elem_str))
            elsif elem_str =~ /^".*"$/
              # String
              str_value = elem_str[1..-2]? || elem_str
              elements << AST::StringLiteral.new(str_value)
            elsif elem_str.upcase == "TRUE"
              elements << AST::Boolean.new(true)
            elsif elem_str.upcase == "FALSE"
              elements << AST::Boolean.new(false)
            else
              # Treat as cell reference or named range
              elements << AST::CellRef.new(elem_str)
            end
          end
        end

        # If no elements were parsed, use a placeholder
        if elements.empty?
          elements << AST::CellRef.new("")
        end

        AST::ArrayConstant.new(elements.map(&.as(AST::Node)))
      when Tokens::EmptyOperand
        AST::CellRef.new("")
      else
        # For other operand types, create a placeholder
        AST::CellRef.new(token.name)
      end
    end
  end
end
