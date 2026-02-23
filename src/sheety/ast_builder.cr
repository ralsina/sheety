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

    def size : Int32
      @node_stack.size
    end

    def [](index : Int32) : AST::Node
      @node_stack[index]
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

    # Convert the AST to a string representation
    def to_s(io : IO) : Nil
      root.to_s(io)
    end

    # Get the AST as a tree structure (for inspection/debugging)
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      root.to_tree(io, indent)
    end

    private def operand_to_ast_node(token : Tokens::Operand) : AST::Node
      case token
      when Tokens::Number
        value = token.compile.as(Float64)
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
      else
        # For other operand types (ranges, references), create a placeholder
        # This will be implemented when we add those features
        AST::CellRef.new(token.name)
      end
    end
  end
end

# Extend AST::Node with tree printing
module Sheety::AST
  class Node
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      io << "  " * indent
      inspect(io)
      io << "\n"
    end
  end

  class UnaryOp
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      io << "  " * indent
      io << "#{operator} <UnaryOp>\n"
      @operand.to_tree(io, indent + 1)
    end
  end

  class BinaryOp
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      io << "  " * indent
      io << "#{operator} <BinaryOp>\n"
      @left.to_tree(io, indent + 1)
      @right.to_tree(io, indent + 1)
    end
  end

  class FunctionCall
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      io << "  " * indent
      io << "#{@function_name} <FunctionCall> (#{@arguments.size} args)\n"
      @arguments.each do |arg|
        arg.to_tree(io, indent + 1)
      end
    end
  end

  class ArrayConstant
    def to_tree(io : IO, indent : Int32 = 0) : Nil
      io << "  " * indent
      io << "<ArrayConstant> (#{@elements.size} elements)\n"
      @elements.each do |elem|
        elem.to_tree(io, indent + 1)
      end
    end
  end
end
