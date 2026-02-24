module Sheety
  # Stores formula information with its sheet context
  private class FormulaInfo
    getter ast : AST::Node
    getter sheet : String?

    def initialize(@ast : AST::Node, @sheet : String? = nil)
    end
  end

  # Evaluates Excel formulas with cell values from a hash
  class Evaluator
    include AST

    @cells : Hash(String, Float64 | String | Bool | Functions::ErrorValue | Nil)
    @formulas : Hash(String, FormulaInfo)
    @current_sheet : String?

    def initialize
      @cells = Hash(String, Float64 | String | Bool | Functions::ErrorValue | Nil).new
      @formulas = Hash(String, FormulaInfo).new
      @current_sheet = nil
    end

    # Set a cell value
    def set(cell : String, value, sheet : String? = nil) : Nil
      key = sheet ? "#{sheet}!#{cell}" : cell
      @cells[key] = value
    end

    # Set a cell formula
    def set_formula(cell : String, formula : String, sheet : String? = nil) : Nil
      key = sheet ? "#{sheet}!#{cell}" : cell
      formula_str = formula.starts_with?("=") ? formula : "=#{formula}"
      ast = Parser.new.ast(formula_str)[1].root
      @formulas[key] = FormulaInfo.new(ast, sheet)
    end

    # Get a cell value, calculating if necessary
    def get(cell : String, sheet : String? = nil)
      key = sheet ? "#{sheet}!#{cell}" : cell

      # If already calculated, return cached value
      return @cells[key] if @cells.has_key?(key)

      # If it's a formula, evaluate it
      if @formulas.has_key?(key)
        formula_info = @formulas[key]
        old_sheet = @current_sheet
        @current_sheet = formula_info.sheet
        result = evaluate(formula_info.ast)
        @current_sheet = old_sheet

        # Convert arrays to something storable
        if result.is_a?(Array)
          # For arrays, we typically don't store them directly
          # They're intermediate values for functions
          return result
        end
        @cells[key] = result
        return result
      end

      # No value or formula
      nil
    end

    # Evaluate an AST node
    def evaluate(node : Node)
      visit(node)
    end

    # Calculate all formulas in dependency order
    def calculate_all : Nil
      changed = true
      iterations = 0
      max_iterations = @formulas.size * 2 # Prevent infinite loops

      while changed && iterations < max_iterations
        changed = false
        iterations += 1

        @formulas.each do |key, formula_info|
          next if @cells.has_key?(key) # Already calculated

          begin
            old_sheet = @current_sheet
            @current_sheet = formula_info.sheet
            result = evaluate(formula_info.ast)
            @current_sheet = old_sheet

            # Skip storing arrays
            unless result.is_a?(Array)
              @cells[key] = result
            end
            changed = true
          rescue e : ArgumentError
            # Missing dependency, try again next iteration
          end
        end
      end
    end

    # Get all cells
    def all_cells
      calculate_all
      @cells.dup
    end

    # Get cells for a specific sheet
    def sheet_cells(sheet_name : String)
      result = Hash(String, Float64 | String | Bool | Functions::ErrorValue | Nil).new
      all_cells.each do |key, value|
        if key.starts_with?("#{sheet_name}!")
          cell = key.sub("#{sheet_name}!", "")
          result[cell] = value
        elsif !key.includes?("!")
          result[key] = value
        end
      end
      result
    end

    # Visitor methods for evaluation

    private def visit(node : Number)
      node.value
    end

    private def visit(node : StringLiteral)
      node.value
    end

    private def visit(node : Boolean)
      node.value
    end

    private def visit(node : ErrorValue)
      Functions::ErrorValue.new(node.error_value)
    end

    private def visit(node : CellRef)
      ref = node.reference
      sheet = node.sheet || @current_sheet
      get(ref, sheet)
    end

    private def visit(node : RangeRef)
      range = node.range
      sheet = node.sheet || @current_sheet

      if match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
        start_col = match[1]
        start_row = match[2].to_i
        end_col = match[3]
        end_row = match[4].to_i

        # Collect all cell values in range and return as array for function processing
        values = expand_range_values(start_col, start_row, end_col, end_row, sheet)
        # For simple ranges used in functions, we need to return them as a flat array
        values
      else
        [] of Functions::CellValue
      end
    end

    private def visit(node : NamedRef)
      Functions::ErrorValue.new("#NAME?")
    end

    private def visit(node : UnaryOp)
      operand = visit(node.operand)

      case node.operator
      when "+"
        to_number(operand)
      when "-"
        num = to_number(operand)
        return num if num.nil?
        -num
      when "%"
        num = to_number(operand)
        return num if num.nil?
        num / 100.0
      else
        nil
      end
    end

    private def visit(node : BinaryOp)
      left = visit(node.left)
      right = visit(node.right)

      case node.operator
      when "+"
        ln = to_number(left)
        rn = to_number(right)
        return nil if ln.nil? || rn.nil?
        ln + rn
      when "-"
        ln = to_number(left)
        rn = to_number(right)
        return nil if ln.nil? || rn.nil?
        ln - rn
      when "*"
        ln = to_number(left)
        rn = to_number(right)
        return nil if ln.nil? || rn.nil?
        ln * rn
      when "/"
        ln = to_number(left)
        rn = to_number(right)
        return nil if ln.nil? || rn.nil?
        return Functions.div0 if rn == 0
        ln / rn
      when "^"
        ln = to_number(left)
        rn = to_number(right)
        return nil if ln.nil? || rn.nil?
        ln ** rn
      when "="
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.eq(left, right)
      when "<>"
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.ne(left, right)
      when "<"
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.lt(left, right)
      when ">"
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.gt(left, right)
      when "<="
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.le(left, right)
      when ">="
        return nil if left.is_a?(Array) || right.is_a?(Array)
        Functions.ge(left, right)
      when "&"
        return "" if left.is_a?(Array) || right.is_a?(Array)
        Functions.to_string(left) + Functions.to_string(right)
      else
        nil
      end
    end

    private def visit(node : FunctionCall)
      # Evaluate arguments and collect all values
      all_values = [] of Functions::CellValue

      node.arguments.each do |arg|
        val = visit(arg)
        if val.is_a?(Array)
          all_values.concat(val)
        else
          all_values << val
        end
      end

      func_name = node.function_name.upcase

      case func_name
      when "SUM"
        Functions.sum(all_values)
      when "AVERAGE", "AVG"
        Functions.average(all_values)
      when "MIN"
        Functions.min(all_values)
      when "MAX"
        Functions.max(all_values)
      when "COUNT"
        Functions.count(all_values)
      when "ROUND"
        Functions.round(all_values[0]?, all_values[1]? || 0.0)
      when "ABS"
        Functions.abs(all_values[0]?)
      when "POWER"
        Functions.power(all_values[0]?, all_values[1]?)
      when "SQRT"
        Functions.sqrt(all_values[0]?)
      when "MOD"
        Functions.mod(all_values[0]?, all_values[1]?)
      when "INT"
        Functions.int(all_values[0]?)
      when "IF"
        Functions.if(all_values[0]?, all_values[1]?, all_values[2]?)
      when "AND"
        Functions.and(all_values)
      when "OR"
        Functions.or(all_values)
      when "NOT"
        Functions.not(all_values[0]?)
      when "CONCAT", "CONCATENATE"
        Functions.concat(all_values)
      when "LEFT"
        Functions.left(all_values[0]?, all_values[1]? || 1.0)
      when "RIGHT"
        Functions.right(all_values[0]?, all_values[1]? || 1.0)
      when "MID"
        Functions.mid(all_values[0]?, all_values[1]? || 1.0, all_values[2]? || 0.0)
      when "LEN"
        Functions.len(all_values[0]?)
      when "UPPER"
        Functions.upper(all_values[0]?)
      when "LOWER"
        Functions.lower(all_values[0]?)
      when "TRIM"
        Functions.trim(all_values[0]?)
      else
        Functions::ErrorValue.new("#NAME?")
      end
    end

    private def visit(node : ArrayConstant)
      result = [] of Functions::CellValue
      node.elements.each do |elem|
        val = visit(elem)
        if val.is_a?(Array)
          result.concat(val)
        else
          result << val
        end
      end
      result
    end

    # Helper to expand range into array of values
    private def expand_range_values(start_col : String, start_row : Int32, end_col : String, end_row : Int32, sheet : String?)
      result = [] of Functions::CellValue

      start_col_num = column_to_number(start_col)
      end_col_num = column_to_number(end_col)

      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = number_to_column(col)
          cell_ref = "#{col_str}#{row}"
          value = get(cell_ref, sheet)
          if value.is_a?(Array)
            result.concat(value)
          elsif !value.nil?
            result << value
          end
        end
      end

      result
    end

    # Helper to convert value to number
    private def to_number(value) : Float64?
      case value
      when Float64 then value
      when String
        begin
          value.to_f
        rescue
          nil
        end
      when Bool  then value ? 1.0 : 0.0
      when Array then nil # Arrays can't be converted to numbers
      else            nil
      end
    end

    # Convert column letter(s) to number (A=1, Z=26, AA=27, etc.)
    private def column_to_number(col : String) : Int32
      num = 0
      col.each_char do |char|
        num = num * 26 + (char.ord - 'A'.ord + 1)
      end
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
