module Sheety
  # Generates Crystal code from Excel formula AST
  #
  # The generated code is designed to be used in procs that:
  # 1. Read cell values from Croupier's k/v store
  # 2. Calculate the formula result
  # 3. Return the result as a string for storage
  class CodeGenerator
    include AST

    # Context for code generation - tracks sheet name and available cells
    class Context
      property sheet : String?
      property cells : Hash(String, Float64 | String | Bool)

      def initialize(@sheet : String? = nil)
        @cells = Hash(String, Float64 | String | Bool).new
      end
    end

    # Generate Crystal code for an AST node
    def generate(node : Node, context : Context = Context.new) : String
      visit(node, context)
    end

    # Generate a complete proc body for a formula
    def generate_proc_body(formula : String, context : Context = Context.new) : String
      ast = Parser.new.ast(formula)[1].root
      code = generate(ast, context)

      # The proc reads from k/v store and returns result as string
      %{
        result = (#{code})

        # Convert result to string for storage
        case result
        when Float64
          # Format numbers appropriately
          if result == result.to_i
            result.to_i.to_s
          else
            result.to_s
          end
        when String
          result
        when Bool
          result.upcase.to_s
        when Sheety::Functions::ErrorValue
          result.to_s
        when Nil
          ""
        else
          result.to_s
        end
      }
    end

    # Visitor methods for each node type

    private def visit(node : Number, context : Context) : String
      node.value.to_s
    end

    private def visit(node : StringLiteral, context : Context) : String
      node.value.inspect
    end

    private def visit(node : Boolean, context : Context) : String
      node.value.to_s
    end

    private def visit(node : ErrorValue, context : Context) : String
      "Sheety::Functions::ErrorValue.new(#{node.error_value.inspect})"
    end

    private def visit(node : CellRef, context : Context) : String
      ref = node.reference
      sheet = node.sheet || context.sheet

      # Generate code to fetch from k/v store with proper nil handling
      key = sheet ? "#{sheet}!#{ref}" : ref
      "(Croupier::TaskManager.get(#{key.inspect}) || \"\")"
    end

    private def visit(node : RangeRef, context : Context) : String
      # Parse and expand range into individual cell references
      range = node.range
      sheet = node.sheet || context.sheet

      # Parse range (e.g., "A1:B5")
      if match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
        start_col = match[1]
        start_row = match[2].to_i
        end_col = match[3]
        end_row = match[4].to_i

        # Generate array of cell fetches with proper nil handling
        cells = expand_range(start_col, start_row, end_col, end_row, sheet)
        "[" + cells.map { |ref| "(Croupier::TaskManager.get(#{ref.inspect}) || \"\")" }.join(", ") + "]"
      else
        "[]"
      end
    end

    private def visit(node : NamedRef, context : Context) : String
      # Named references would need to be resolved beforehand
      # For now, return an error
      "Sheety::Functions::ErrorValue.new(\"#NAME?\")"
    end

    private def visit(node : UnaryOp, context : Context) : String
      operand = visit(node.operand, context)

      case node.operator
      when "+"
        "(+(#{operand}))"
      when "-"
        "(-(#{operand}))"
      when "%"
        "((#{operand}) / 100.0)"
      else
        "(#{node.operator} #{operand})"
      end
    end

    private def visit(node : BinaryOp, context : Context) : String
      left = visit(node.left, context)
      right = visit(node.right, context)

      case node.operator
      when "="
        "Sheety::Functions.eq(#{left}, #{right})"
      when "<>"
        "Sheety::Functions.ne(#{left}, #{right})"
      when "<"
        "Sheety::Functions.lt(#{left}, #{right})"
      when ">"
        "Sheety::Functions.gt(#{left}, #{right})"
      when "<="
        "Sheety::Functions.le(#{left}, #{right})"
      when ">="
        "Sheety::Functions.ge(#{left}, #{right})"
      when "&"
        "Sheety::Functions.to_string(#{left}) + Sheety::Functions.to_string(#{right})"
      when "+"
        "begin; ln = Sheety::Functions.to_float(#{left}); rn = Sheety::Functions.to_float(#{right}); (ln && rn) ? (ln + rn).to_s : nil; end"
      when "-"
        "begin; ln = Sheety::Functions.to_float(#{left}); rn = Sheety::Functions.to_float(#{right}); (ln && rn) ? (ln - rn).to_s : nil; end"
      when "*"
        "begin; ln = Sheety::Functions.to_float(#{left}); rn = Sheety::Functions.to_float(#{right}); (ln && rn) ? (ln * rn).to_s : nil; end"
      when "/"
        "begin; ln = Sheety::Functions.to_float(#{left}); rn = Sheety::Functions.to_float(#{right}); (ln && rn) ? (ln / rn).to_s : nil; end"
      when "^"
        "begin; ln = Sheety::Functions.to_float(#{left}); rn = Sheety::Functions.to_float(#{right}); (ln && rn) ? (ln ** rn).to_s : nil; end"
      else
        "((#{left}) #{node.operator} (#{right}))"
      end
    end

    private def visit(node : FunctionCall, context : Context) : String
      func_name = node.function_name.upcase
      args = node.arguments.map { |arg| visit(arg, context) }

      case func_name
      when "SUM"
        "Sheety::Functions.sum(#{args.join(", ")})"
      when "AVERAGE", "AVG"
        "Sheety::Functions.average(#{args.join(", ")})"
      when "MIN"
        "Sheety::Functions.min(#{args.join(", ")})"
      when "MAX"
        "Sheety::Functions.max(#{args.join(", ")})"
      when "COUNT"
        "Sheety::Functions.count(#{args.join(", ")})"
      when "ROUND"
        "Sheety::Functions.round(#{args.join(", ")})"
      when "ABS"
        "Sheety::Functions.abs(#{args[0]})"
      when "POWER"
        "Sheety::Functions.power(#{args.join(", ")})"
      when "SQRT"
        "Sheety::Functions.sqrt(#{args[0]})"
      when "MOD"
        "Sheety::Functions.mod(#{args.join(", ")})"
      when "INT"
        "Sheety::Functions.int(#{args[0]})"
      when "IF"
        if args.size >= 3
          "Sheety::Functions.if(#{args[0]}, #{args[1]}, #{args[2]})"
        else
          "Sheety::Functions::ErrorValue.new(\"#VALUE!\")"
        end
      when "AND"
        "Sheety::Functions.and([#{args.join(", ")}])"
      when "OR"
        "Sheety::Functions.or([#{args.join(", ")}])"
      when "NOT"
        "Sheety::Functions.not(#{args[0]})"
      when "CONCAT", "CONCATENATE"
        "Sheety::Functions.concat([#{args.join(", ")}])"
      when "LEFT"
        "Sheety::Functions.left(#{args.join(", ")})"
      when "RIGHT"
        "Sheety::Functions.right(#{args.join(", ")})"
      when "MID"
        if args.size >= 3
          "Sheety::Functions.mid(#{args.join(", ")})"
        else
          "Sheety::Functions::ErrorValue.new(\"#VALUE!\")"
        end
      when "LEN"
        "Sheety::Functions.len(#{args[0]})"
      when "UPPER"
        "Sheety::Functions.upper(#{args[0]})"
      when "LOWER"
        "Sheety::Functions.lower(#{args[0]})"
      when "TRIM"
        "Sheety::Functions.trim(#{args[0]})"
      else
        # Unknown function
        "Sheety::Functions::ErrorValue.new(\"#NAME?\")"
      end
    end

    private def visit(node : ArrayConstant, context : Context) : String
      elements = node.elements.map { |elem| visit(elem, context) }
      "[#{elements.join(", ")}]"
    end

    # Helper to expand a range like "A1:B2" into cell references
    private def expand_range(start_col : String, start_row : Int32, end_col : String, end_row : Int32, sheet : String?) : Array(String)
      result = [] of String

      # Convert column letters to numbers
      start_col_num = column_to_number(start_col)
      end_col_num = column_to_number(end_col)

      # Iterate through rows and columns
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = number_to_column(col)
          ref = sheet ? "#{sheet}!#{col_str}#{row}" : "#{col_str}#{row}"
          result << ref
        end
      end

      result
    end

    # Convert column letter(s) to number (A=1, Z=26, AA=27, etc.)
    private def column_to_number(col : String) : Int32
      num = 0
      col.each_char { |char| num = num * 26 + (char.ord - 'A'.ord + 1) }
      num
    end

    # Convert column number to letter(s) (1=A, 26=Z, 27=AA, etc.)
    private def number_to_column(num : Int32) : String
      result = ""
      while num > 0
        num -= 1
        result = ('A' + (num % 26)).to_s + result
        num //= 26
      end
      result
    end
  end
end
