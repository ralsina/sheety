require "../functions/registry"

module Sheety
  module Importers
    # Represents a single cell from an Excel file
    # Contains both the calculated value and the formula (if any)
    struct ExcelCell
      property reference : String
      property value : Functions::CellValue?
      property formula : String?

      def initialize(@reference : String, @value : Functions::CellValue? = nil, @formula : String? = nil)
      end
    end

    # Represents a sheet from an Excel file
    # Contains the sheet name and all cells in that sheet
    struct ExcelSheet
      property name : String
      property cells : Array(ExcelCell)

      def initialize(@name : String, @cells : Array(ExcelCell) = Array(ExcelCell).new)
      end

      # Get cell by reference
      def cell(reference : String) : ExcelCell?
        @cells.find { |cell| cell.reference == reference }
      end

      # Get all cell references sorted
      def sorted_references : Array(String)
        @cells.map(&.reference).sort! do |ref_a, ref_b|
          ref_a_parts = parse_reference(ref_a)
          ref_b_parts = parse_reference(ref_b)

          # Compare by column first, then by row
          if ref_a_parts[0] == ref_b_parts[0]
            ref_a_parts[1] <=> ref_b_parts[1]
          else
            ref_a_parts[0] <=> ref_b_parts[0]
          end
        end
      end

      # Parse cell reference into (column, row_number) tuple
      # e.g., "A1" => {"A", 1}, "Z10" => {"Z", 10}
      private def parse_reference(ref : String) : Tuple(String, Int32)
        match = ref.match(/^([A-Z]+)(\d+)$/)
        if match
          {match[1], match[2].to_i}
        else
          {"", 0}
        end
      end
    end

    # Represents a complete Excel workbook
    struct ExcelWorkbook
      property sheets : Array(ExcelSheet)

      def initialize(@sheets : Array(ExcelSheet) = Array(ExcelSheet).new)
      end

      # Get sheet by name
      def sheet(name : String) : ExcelSheet?
        @sheets.find { |sheet| sheet.name == name }
      end

      # Get all sheet names
      def sheet_names : Array(String)
        @sheets.map(&.name)
      end
    end
  end
end
