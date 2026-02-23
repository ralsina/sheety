module Sheety
  module AST
    # Base class for all AST nodes
    abstract class Node
      property attr : Hash(String, String | Bool | Float64 | Int32)

      def initialize
        @attr = Hash(String, String | Bool | Float64 | Int32).new
      end

      # Get the expression string for this node
      def expr : String
        @attr.fetch("expr", "").as(String)
      end

      def to_s(io : IO) : Nil
        io << expr
      end

      def inspect(io : IO) : Nil
        io << "<#{self.class.name} #{expr}>"
      end
    end

    # Number literal node
    class Number < Node
      property value : Float64

      def initialize(@value : Float64)
        super()
        @attr["value"] = @value
        @attr["expr"] = format_number(@value)
      end

      def expr : String
        format_number(@value)
      end

      private def format_number(value : Float64) : String
        if value == value.to_i
          value.to_i.to_s
        else
          value.to_s
        end
      end
    end

    # String literal node
    class StringLiteral < Node
      property value : String

      def initialize(@value : String)
        super()
        @attr["value"] = @value
        @attr["expr"] = "\"#{@value}\""
      end

      def expr : String
        "\"#{@value}\""
      end
    end

    # Boolean literal node
    class Boolean < Node
      property value : Bool

      def initialize(@value : Bool)
        super()
        @attr["value"] = @value ? 1.0 : 0.0
        @attr["expr"] = @value.to_s.upcase
      end

      def expr : String
        @value.to_s.upcase
      end
    end

    # Error value node
    class ErrorValue < Node
      property error_value : String

      def initialize(@error_value : String)
        super()
        @attr["error"] = @error_value
        @attr["expr"] = @error_value
      end

      def expr : String
        @error_value
      end
    end

    # Cell reference node (e.g., A1, Sheet1!B5)
    class CellRef < Node
      property reference : String

      def initialize(@reference : String)
        super()
        @attr["reference"] = @reference
        @attr["expr"] = @reference
      end

      def expr : String
        @reference
      end
    end

    # Range reference node (e.g., A1:B5)
    class RangeRef < Node
      property range : String

      def initialize(@range : String)
        super()
        @attr["range"] = @range
        @attr["expr"] = @range
      end

      def expr : String
        @range
      end
    end

    # Unary operation node (e.g., -5, +3)
    class UnaryOp < Node
      property operator : String
      property operand : Node

      def initialize(@operator : String, @operand : Node)
        super()
        @attr["operator"] = @operator
        @attr["expr"] = "(#{@operator}#{@operand.expr})"
      end

      def expr : String
        "(#{@operator}#{@operand.expr})"
      end
    end

    # Binary operation node (e.g., 1 + 2, A1 * B1)
    class BinaryOp < Node
      property operator : String
      property left : Node
      property right : Node

      def initialize(@operator : String, @left : Node, @right : Node)
        super()
        @attr["operator"] = @operator
        @attr["expr"] = "(#{@left.expr} #{@operator} #{@right.expr})"
      end

      def expr : String
        "(#{@left.expr} #{@operator} #{@right.expr})"
      end
    end

    # Function call node (e.g., SUM(A1:B5), IF(A1>0, 1, 0))
    class FunctionCall < Node
      property function_name : String
      property arguments : Array(Node)

      def initialize(@function_name : String, @arguments : Array(Node) = Array(Node).new)
        super()
        @attr["function"] = @function_name
        args_str = @arguments.map(&.expr).join(", ")
        @attr["expr"] = "#{@function_name}(#{args_str})"
      end

      def expr : String
        args_str = @arguments.map(&.expr).join(", ")
        "#{@function_name}(#{args_str})"
      end
    end

    # Array constant node (e.g., {1, 2, 3})
    class ArrayConstant < Node
      property elements : Array(Node)

      def initialize(@elements : Array(Node))
        super()
        @attr["elements"] = @elements
        @attr["expr"] = "{#{@elements.map(&.expr).join(", ")}}"
      end

      def expr : String
        "{#{@elements.map(&.expr).join(", ")}}"
      end
    end
  end
end
