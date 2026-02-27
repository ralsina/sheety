require "xlsx-parser"
require "./excel_types"
require "./formula_extractor"

module Sheety
  class ExcelImporter
    # Parse an Excel (.xlsx) file and return an ExcelWorkbook
    def self.parse_xlsx(filename : String) : Importers::ExcelWorkbook
      unless File.exists?(filename)
        raise "File not found: #{filename}"
      end

      # Check file extension
      unless filename.downcase.ends_with?(".xlsx")
        raise "Invalid file format. Expected .xlsx file, got: #{File.extname(filename)}"
      end

      begin
        book = XlsxParser::Book.new(filename)
      rescue ex : Exception
        raise "Unable to open Excel file: #{ex.message}"
      end

      begin
        sheets = Array(Importers::ExcelSheet).new

        book.sheets.each_with_index do |xlsx_sheet, index|
          # Extract values using xlsx-parser
          values = extract_values_from_sheet(xlsx_sheet)

          # Extract formulas from XML
          formulas = Sheety::FormulaExtractor.extract(filename, index)

          # Merge into ExcelCell structures
          cells = merge_values_and_formulas(values, formulas)

          sheets << Importers::ExcelSheet.new(xlsx_sheet.name, cells)
        end

        Importers::ExcelWorkbook.new(sheets)
      ensure
        book.close
      end
    end

    # Convert ExcelWorkbook to Sheety's internal format
    # Returns: Hash(String, Hash(String, Hash(String, Functions::CellValue)))
    # Compatible with the YAML structure expected by CroupierGenerator
    def self.to_internal_format(workbook : Importers::ExcelWorkbook) : Hash(String, Hash(String, Hash(String, Functions::CellValue)))
      result = {} of String => Hash(String, Hash(String, Functions::CellValue))

      workbook.sheets.each do |sheet|
        sheet_data = {} of String => Hash(String, Functions::CellValue)

        sheet.cells.each do |cell|
          cell_data = {} of String => Functions::CellValue

          if cell.formula
            # Store formula as string (Sheety will parse it)
            cell_data["formula"] = cell.formula
            sheet_data[cell.reference] = cell_data
          elsif !cell.value.nil?
            # Store the value directly (including false, 0, empty string, etc.)
            cell_data["value"] = cell.value
            sheet_data[cell.reference] = cell_data
          end
        end

        result[sheet.name] = sheet_data
      end

      result
    end

    # Extract cell values from xlsx-parser sheet
    private def self.extract_values_from_sheet(xlsx_sheet : XlsxParser::Sheet) : Hash(String, Functions::CellValue)
      values = {} of String => Functions::CellValue

      # xlsx-parser provides rows as Hash(String, Type) where the key is cell reference (A1, B1, etc.)
      # and the value is the cell value
      xlsx_sheet.rows.each do |row|
        row.each do |cell_ref, cell_value|
          # Skip nil values
          next if cell_value.nil?

          # Convert the value to Sheety's CellValue type
          values[cell_ref] = convert_value(cell_value)
        end
      end

      values
    end

    # Merge values and formulas into ExcelCell objects
    private def self.merge_values_and_formulas(
      values : Hash(String, Functions::CellValue),
      formulas : Hash(String, String),
    ) : Array(Importers::ExcelCell)
      cells = Array(Importers::ExcelCell).new
      processed_refs = Set(String).new

      # First, process cells that have formulas
      formulas.each do |ref, formula|
        value = values[ref]?
        cells << Importers::ExcelCell.new(ref, value, formula)
        processed_refs << ref
      end

      # Then, add cells that only have values
      values.each do |ref, value|
        unless processed_refs.includes?(ref)
          cells << Importers::ExcelCell.new(ref, value, nil)
        end
      end

      cells
    end

    # Convert xlsx-parser value to Sheety's CellValue type
    private def self.convert_value(value) : Functions::CellValue
      case value
      when Int32, Int64
        value.to_f
      when Float64
        value
      when String
        value
      when Bool
        value
      when Time
        # For now, store dates as their serial number
        # Excel stores dates as days since 1900-01-01
        epoch = Time.utc(1900, 1, 1)
        seconds = (value - epoch).total_seconds.to_i64
        days = (seconds // 86400).to_i
        (days + 2).to_f # Excel's 1900 date system has a bug treating 1900 as leap year
      when Nil
        nil
      else
        # Fallback: convert to string
        value.to_s
      end
    end
  end
end
