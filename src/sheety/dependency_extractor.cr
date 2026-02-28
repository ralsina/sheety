module Sheety
  # Extracts cell dependencies from formula AST for Croupier task inputs
  class DependencyExtractor
    include AST

    # Extract cell references from an AST node
    def extract(node : Node, sheet : String? = nil) : Set(String)
      dependencies = Set(String).new
      visit(node, dependencies, sheet)
      dependencies
    end

    # Extract from formula string
    def extract_from_formula(formula : String, sheet : String? = nil) : Set(String)
      ast = Parser.new.ast(formula)[1].root
      extract(ast, sheet)
    end

    # Visitor methods for each node type

    private def visit(node : Number, dependencies : Set(String), sheet : String?) : Nil
      # Numbers have no dependencies
    end

    private def visit(node : StringLiteral, dependencies : Set(String), sheet : String?) : Nil
      # Strings have no dependencies
    end

    private def visit(node : Boolean, dependencies : Set(String), sheet : String?) : Nil
      # Booleans have no dependencies
    end

    private def visit(node : ErrorValue, dependencies : Set(String), sheet : String?) : Nil
      # Errors have no dependencies
    end

    private def visit(node : CellRef, dependencies : Set(String), sheet : String?) : Nil
      ref = node.reference.upcase
      cell_sheet = node.sheet || sheet
      key = cell_sheet ? "#{cell_sheet}!#{ref}" : ref
      dependencies.add(key)
    end

    private def visit(node : RangeRef, dependencies : Set(String), sheet : String?) : Nil
      # Expand range into individual cell references
      range = node.range.upcase
      cell_sheet = node.sheet || sheet

      if match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
        start_col = match[1]
        start_row = match[2].to_i
        end_col = match[3]
        end_row = match[4].to_i

        # Add each cell in range as a dependency
        expand_range(start_col, start_row, end_col, end_row, cell_sheet).each do |ref|
          dependencies.add(ref)
        end
      end
    end

    private def visit(node : NamedRef, dependencies : Set(String), sheet : String?) : Nil
      # Named references would need external resolution
      # For now, we don't track them as dependencies
    end

    private def visit(node : UnaryOp, dependencies : Set(String), sheet : String?) : Nil
      visit(node.operand, dependencies, sheet)
    end

    private def visit(node : BinaryOp, dependencies : Set(String), sheet : String?) : Nil
      visit(node.left, dependencies, sheet)
      visit(node.right, dependencies, sheet)
    end

    private def visit(node : FunctionCall, dependencies : Set(String), sheet : String?) : Nil
      node.arguments.each do |arg|
        visit(arg, dependencies, sheet)
      end
    end

    private def visit(node : ArrayConstant, dependencies : Set(String), sheet : String?) : Nil
      node.elements.each do |elem|
        visit(elem, dependencies, sheet)
      end
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
