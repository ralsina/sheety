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

    # Fetch a range of cells (e.g., "Sheet1!A1:A100")
    def fetch_cell_range(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Array(String)
      # Convert column letters to numbers
      col_num = ->(col : String) {
        num = 0
        col.each_char { |char| num = num * 26 + (char.ord - 'A'.ord + 1) }
        num
      }

      # Convert column numbers to letters
      num_to_col = ->(num : Int32) {
        result = ""
        while num > 0
          num -= 1
          result = ('A' + (num % 26)).to_s + result
          num //= 26
        end
        result
      }

      start_col_num = col_num.call(start_col)
      end_col_num = col_num.call(end_col)

      # Build array of cell references and fetch values
      result = [] of String
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = num_to_col.call(col)
          result << fetch_cell(sheet + "!" + col_str + row.to_s)
        end
      end
      result
    end

    # Generate k/v store input keys for a range
    def range_inputs(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Array(String)
      # Convert column letters to numbers
      col_num = ->(col : String) {
        num = 0
        col.each_char { |char| num = num * 26 + (char.ord - 'A'.ord + 1) }
        num
      }

      # Convert column numbers to letters
      num_to_col = ->(num : Int32) {
        result = ""
        while num > 0
          num -= 1
          result = ('A' + (num % 26)).to_s + result
          num //= 26
        end
        result
      }

      start_col_num = col_num.call(start_col)
      end_col_num = col_num.call(end_col)

      # Build array of k/v store keys
      result = [] of String
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = num_to_col.call(col)
          result << "kv://" + sheet + "!" + col_str + row.to_s
        end
      end
      result
    end

    # Convert task result to string for k/v store output
    def format_result(result) : String
      case result
      when Float64
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
    end

    # Initialize multiple cells at once from a hash
    def initialize_cells(cells : Hash(String, String))
      cells.each do |key, value|
        Croupier::TaskManager.set(key, value)
      end
    end

    # Initialize all cells in a range to empty strings
    # This is needed because Croupier requires all input keys to exist
    def initialize_range(sheet : String, start_col : String, start_row : Int32, end_col : String, end_row : Int32) : Nil
      # Convert column letters to numbers
      col_num = ->(col : String) {
        num = 0
        col.each_char { |char| num = num * 26 + (char.ord - 'A'.ord + 1) }
        num
      }

      # Convert column numbers to letters
      num_to_col = ->(num : Int32) {
        result = ""
        while num > 0
          num -= 1
          result = ('A' + (num % 26)).to_s + result
          num //= 26
        end
        result
      }

      start_col_num = col_num.call(start_col)
      end_col_num = col_num.call(end_col)

      # Set all cells in range to empty string
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = num_to_col.call(col)
          Croupier::TaskManager.set(sheet + "!" + col_str + row.to_s, "")
        end
      end
    end
  end
end

# Include the helpers at the top level so they're available in the generated code
include Sheety::CroupierHelpers
