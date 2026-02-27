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
        col.each_char { |c| num = num * 26 + (c.ord - 'A'.ord + 1) }
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
  end
end

# Include the helpers at the top level so they're available in the generated code
include Sheety::CroupierHelpers
