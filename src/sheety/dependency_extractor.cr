require "./cell_refs"
require "./errors"

module Sheety
  # Extracts cell dependencies from formula AST for Croupier task inputs
  class DependencyExtractor
    include AST

    # A normalized range reference visited while extracting dependencies: the
    # sheet it resolved to plus bounds already normalized by
    # CellRefs.parse_range. The generator emits its initialize_range calls
    # from these instead of regex-scanning its own generated code, so the
    # two can never disagree.
    record RangeDependency, sheet : String?, bounds : CellRefs::RangeBounds

    # What #extract_with_ranges collected: the per-cell dependency keys and
    # the range bounds they were expanded from.
    record Extraction, dependencies : Set(String), ranges : Set(RangeDependency)

    # Sanity cap on expanded range size. A typo like A1:B99999999 would
    # otherwise expand into billions of dependency keys (and just as many
    # kv entries at runtime) before anything could stop it. Formulas using
    # oversized ranges raise FormulaError; the generator turns those into
    # #VALUE! tasks instead of hanging or exhausting memory.
    MAX_RANGE_CELLS = 65536

    # Extract cell references from an AST node
    def extract(node : Node, sheet : String? = nil) : Set(String)
      extract_with_ranges(node, sheet).dependencies
    end

    # Extract from formula string
    def extract_from_formula(formula : String, sheet : String? = nil) : Set(String)
      ast = Parser.new.ast(formula)[1].root
      extract(ast, sheet)
    end

    # Extract cell dependencies and the normalized range bounds they were
    # expanded from, in a single AST walk so the two results always agree.
    def extract_with_ranges(node : Node, sheet : String? = nil) : Extraction
      dependencies = Set(String).new
      ranges = Set(RangeDependency).new
      visit(node, dependencies, ranges, sheet)
      Extraction.new(dependencies, ranges)
    end

    # Visitor methods for each node type

    private def visit(node : Number, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Numbers have no dependencies
    end

    private def visit(node : StringLiteral, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Strings have no dependencies
    end

    private def visit(node : Boolean, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Booleans have no dependencies
    end

    private def visit(node : ErrorValue, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Errors have no dependencies
    end

    private def visit(node : CellRef, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Strip $ anchors so dependencies match the keys the generated code
      # actually fetches.
      ref = node.reference.upcase.delete('$')
      cell_sheet = node.sheet || sheet
      key = cell_sheet ? "#{cell_sheet}!#{ref}" : ref
      dependencies.add(key)
    end

    private def visit(node : RangeRef, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Normalize through CellRefs.parse_range so the extracted dependencies
      # cover exactly the cells the generated code will fetch ($ anchors,
      # reversed bounds, whole-column clamping). Unsupported ranges raise:
      # the generator turns those formulas into loud #VALUE! tasks rather
      # than silently depending on nothing.
      bounds = CellRefs.parse_range(node.range)
      raise FormulaError.new("Unsupported range reference: #{node.range}") unless bounds
      cell_sheet = node.sheet || sheet

      # Record the bounds before expanding: the generator initializes the
      # range from them. Oversized ranges still raise below, and the caller
      # discards this formula's results on that path.
      ranges << RangeDependency.new(cell_sheet, bounds)

      # Add each cell in range as a dependency
      expand_range(bounds.start_col, bounds.start_row, bounds.end_col, bounds.end_row, cell_sheet).each do |ref|
        dependencies.add(ref)
      end
    end

    private def visit(node : NamedRef, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      # Named references would need external resolution
      # For now, we don't track them as dependencies
    end

    private def visit(node : UnaryOp, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      visit(node.operand, dependencies, ranges, sheet)
    end

    private def visit(node : BinaryOp, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      visit(node.left, dependencies, ranges, sheet)
      visit(node.right, dependencies, ranges, sheet)
    end

    private def visit(node : FunctionCall, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      node.arguments.each do |arg|
        visit(arg, dependencies, ranges, sheet)
      end
    end

    private def visit(node : ArrayConstant, dependencies : Set(String), ranges : Set(RangeDependency), sheet : String?) : Nil
      node.elements.each do |elem|
        visit(elem, dependencies, ranges, sheet)
      end
    end

    # Helper to expand a range like "A1:B2" into cell references
    private def expand_range(start_col : String, start_row : Int32, end_col : String, end_row : Int32, sheet : String?) : Array(String)
      # Convert column letters to numbers
      start_col_num = CellRefs.col_to_num(start_col)
      end_col_num = CellRefs.col_to_num(end_col)

      # Reject oversized ranges before expanding (reversed ranges expand to
      # nothing, as they always did).
      rows = end_row - start_row + 1
      columns = end_col_num - start_col_num + 1
      if rows > 0 && columns > 0 && rows * columns > MAX_RANGE_CELLS
        raise FormulaError.new(
          "Range #{start_col}#{start_row}:#{end_col}#{end_row} expands to #{rows * columns} cells " \
          "(limit is #{MAX_RANGE_CELLS})"
        )
      end

      result = [] of String

      # Iterate through rows and columns
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = CellRefs.num_to_col(col)
          ref = sheet ? "#{sheet}!#{col_str}#{row}" : "#{col_str}#{row}"
          result << ref
        end
      end

      result
    end
  end
end
