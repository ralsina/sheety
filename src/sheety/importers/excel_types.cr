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
    end

    # Represents a complete Excel workbook
    struct ExcelWorkbook
      property sheets : Array(ExcelSheet)

      def initialize(@sheets : Array(ExcelSheet) = Array(ExcelSheet).new)
      end
    end
  end
end
