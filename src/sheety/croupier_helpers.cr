require "big"
require "./cell_refs"

module Sheety
  # Helper functions for generated spreadsheet binaries
  #
  # These are included in the generated Crystal code to provide
  # efficient cell value fetching operations.
  module CroupierHelpers
    # Fetch a single cell value with default fallback
    def fetch_cell(cell_ref : String) : String
      Croupier::TaskManager.get(cell_ref) || ""
    end

    # Fetch a range of cells (e.g., "Sheet1!A1:A100") as a flat,
    # row-major array of values.
    def fetch_cell_range(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Array(String)
      start_col_num = CellRefs.col_to_num(start_col)
      end_col_num = CellRefs.col_to_num(end_col)

      # Build array of cell references and fetch values
      result = [] of String
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = CellRefs.num_to_col(col)
          result << fetch_cell(sheet + "!" + col_str + row.to_s)
        end
      end
      result
    end

    # Fetch a range of cells as a 2D table (one Array of values per row),
    # for functions that take a table argument (VLOOKUP, HLOOKUP, INDEX).
    # Takes the same arguments as fetch_cell_range so the generator's range
    # scanning treats both helpers uniformly.
    def fetch_cell_range_2d(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Array(Array(String))
      start_col_num = CellRefs.col_to_num(start_col)
      end_col_num = CellRefs.col_to_num(end_col)

      result = [] of Array(String)
      (start_row..end_row).each do |row|
        current_row = [] of String
        (start_col_num..end_col_num).each do |col|
          col_str = CellRefs.num_to_col(col)
          current_row << fetch_cell(sheet + "!" + col_str + row.to_s)
        end
        result << current_row
      end
      result
    end

    # Generate k/v store input keys for a range
    def range_inputs(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Array(String)
      start_col_num = CellRefs.col_to_num(start_col)
      end_col_num = CellRefs.col_to_num(end_col)

      # Build array of k/v store keys
      result = [] of String
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = CellRefs.num_to_col(col)
          result << "kv://" + sheet + "!" + col_str + row.to_s
        end
      end
      result
    end

    # Convert task result to string for k/v store output
    def format_result(result) : String
      case result
      when BigFloat
        if result == result.to_i
          result.to_i.to_s
        else
          result.to_s
        end
      when String
        result
      when Bool
        result ? "TRUE" : "FALSE"
      when Sheety::Functions::ErrorValue
        result.to_s
      when Nil
        ""
      else
        result.to_s
      end
    end

    # Helper for binary arithmetic operations
    def bin_add(left, right) : String?
      ln = Sheety::Functions.to_float(left)
      rn = Sheety::Functions.to_float(right)
      return nil unless ln && rn
      format_result(ln + rn)
    end

    def bin_sub(left, right) : String?
      ln = Sheety::Functions.to_float(left)
      rn = Sheety::Functions.to_float(right)
      return nil unless ln && rn
      format_result(ln - rn)
    end

    def bin_mul(left, right) : String?
      ln = Sheety::Functions.to_float(left)
      rn = Sheety::Functions.to_float(right)
      return nil unless ln && rn
      format_result(ln * rn)
    end

    def bin_div(left, right) : String?
      ln = Sheety::Functions.to_float(left)
      rn = Sheety::Functions.to_float(right)
      return nil unless ln && rn
      format_result(ln / rn)
    end

    def bin_pow(left, right) : String?
      ln = Sheety::Functions.to_float(left)
      rn = Sheety::Functions.to_float(right)
      return nil unless ln && rn
      # For integer exponents, use BigFloat's ** operator
      if rn == rn.to_i
        format_result(ln ** rn.to_i)
      else
        # For non-integer exponents, convert to Float and use ** operator
        format_result(BigFloat.new(ln.to_f ** rn.to_f, precision: 64))
      end
    end

    # Initialize multiple cells at once from a hash
    def initialize_cells(cells : Hash(String, String))
      cells.each do |key, value|
        Croupier::TaskManager.set(key, value)
      end
    end

    # Register one formula task. This is a thin wrapper around Croupier::Task.new
    # so that the generated source can register thousands of tasks from a single
    # data-driven loop (one Task.new call site in the compiled program) instead
    # of emitting a literal Task.new block per cell — which made the Crystal
    # compiler run out of memory on large sheets.
    def register_formula_task(id : String, inputs : Array(String), output : String, &block : -> String) : Nil
      Croupier::Task.new(
        id: id,
        inputs: inputs,
        outputs: [output],
      ) do
        block.call
      end
    end

    # Initialize all cells in a range to empty strings
    # This is needed because Croupier requires all input keys to exist
    def initialize_range(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Nil
      start_col_num = CellRefs.col_to_num(start_col)
      end_col_num = CellRefs.col_to_num(end_col)

      # Set all cells in range to empty string
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = CellRefs.num_to_col(col)
          Croupier::TaskManager.set(sheet + "!" + col_str + row.to_s, "")
        end
      end
    end
  end
end

# Include the helpers at the top level so they're available in the generated code
include Sheety::CroupierHelpers
