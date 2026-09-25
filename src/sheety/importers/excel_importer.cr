require "big"
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

        # Sheet names come from workbook.xml; cell values are read with eager
        # DOM parsing below instead of XlsxParser::Sheet#rows. The shard
        # walks each sheet with a lazy XML::Reader spread across many
        # allocations, and on loaded CI runners it intermittently yielded
        # sheets with no values at all (the roundtrip spec flake; the zip on
        # disk was verifiably complete). One XML.parse call has no cross-call
        # reader state to lose.
        workbook = XML.parse(book.zip["xl/workbook.xml"].open(&.gets_to_end))
        sheet_nodes = workbook.xpath_nodes("//*[name()='sheet']")

        sheet_nodes.each_with_index do |sheet_node, index|
          sheet_name = sheet_node["name"]? || "Sheet#{index + 1}"

          sheet_path = Sheety::FormulaExtractor.resolve_sheet_path(book.zip, index)
          sheet_xml = sheet_path ? book.zip[sheet_path]?.try(&.open(&.gets_to_end)) : nil
          values = extract_values_from_xml(sheet_xml, book)

          # Extract formulas from XML
          formulas = Sheety::FormulaExtractor.extract(filename, index)

          # Merge into ExcelCell structures
          cells = merge_values_and_formulas(values, formulas)

          sheets << Importers::ExcelSheet.new(sheet_name, cells)
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

    # Style categories that turn an otherwise-numeric cell into a date/time,
    # mirroring XlsxParser::Styles::Converter::DATE_TYPES.
    DATE_TYPES = {:date, :time, :date_time}

    # Extract cell values (and the XML "t" type per cell, for
    # numeric-looking text promotion in convert_value) from one sheet's XML.
    #
    # Eagerly DOM-parses the document rather than delegating to
    # XlsxParser::Sheet#rows; see parse_xlsx for why the lazy reader is
    # avoided. Value typing mirrors XlsxParser::Styles::Converter so the
    # values fed into convert_value keep the exact types the shard produced.
    private def self.extract_values_from_xml(sheet_xml : String?, book : XlsxParser::Book) : Hash(String, Functions::CellValue)
      values = {} of String => Functions::CellValue
      return values unless sheet_xml

      doc = XML.parse(sheet_xml)
      shared_strings = book.shared_strings
      base_time = book.base_time

      doc.xpath_nodes("//*[local-name()='row']/*[local-name()='c']").each do |cell_node|
        cell_ref = cell_node.attributes["r"]?.try(&.content)
        next unless cell_ref

        cell_type = cell_node.attributes["t"]?.try(&.content)

        # Only cells carrying a <v> child hold a (cached) value; like the
        # shard, cells with just a formula or inline content yield nothing.
        v_nodes = cell_node.xpath_nodes("./*[local-name()='v']")
        next if v_nodes.empty?

        style_index = cell_node.attributes["s"]?.try(&.content.try(&.to_i?))
        style = style_index ? book.style_types[style_index]? : nil

        raw = apply_cell_style(v_nodes[v_nodes.size - 1].content, cell_type, style, shared_strings, base_time)
        values[cell_ref] = convert_value(raw, cell_type)
      end

      values
    end

    # Turn a cell's raw <v> text into the type the XML "t" attribute (or the
    # cell style, for date formats) calls for. Replicates
    # XlsxParser::Styles::Converter#call so imports keep producing the same
    # Int32/Int64/Float64/Bool/String/Time values the shard did.
    private def self.apply_cell_style(raw : String, cell_type : String?, style : Symbol?, shared_strings : Array(String), base_time : Time) : Bool | Float64 | Int32 | Int64 | String | Time
      resolved_type = if cell_type.nil? || (cell_type == "n" && DATE_TYPES.includes?(style))
                        style
                      else
                        cell_type
                      end

      case resolved_type
      when "s"
        shared_strings[raw.to_i]
      when "b"
        raw.to_i == 1
      when "n", :float, :percentage
        number = raw.to_f?
        number && number.to_s == raw ? number : raw
      when :fixnum
        int32 = raw.to_i32?
        if int32 && int32.to_s == raw
          int32
        else
          int64 = raw.to_i64?
          int64 && int64.to_s == raw ? int64 : raw
        end
      when :time, :date, :date_time
        base_time + raw.to_f.days
      when :string
        raw
      else
        int32 = raw.to_i32?
        if int32 && int32.to_s == raw
          int32
        else
          number = raw.to_f?
          number && number.to_s == raw ? number : raw
        end
      end
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

    # Convert xlsx-parser value to Sheety's CellValue type. The cell_type is
    # the raw XML t attribute ("n"/absent = numeric, "s" = text, ...):
    # xlsx-parser returns integer-valued numeric cells as strings (its float
    # conversion requires round-tripping through Float64#to_s), so a
    # numeric-looking string in a numeric-typed cell is promoted back to a
    # number while text cells keep their strings.
    private def self.convert_value(value, cell_type : String?) : Functions::CellValue
      if value.is_a?(String) && (cell_type.nil? || cell_type == "n")
        if num = value.to_f?
          return BigFloat.new(num, precision: Functions::DEFAULT_PRECISION)
        end
      end

      case value
      when Int32, Int64
        BigFloat.new(value.to_f, precision: Functions::DEFAULT_PRECISION)
      when BigFloat
        BigFloat.new(value, precision: Functions::DEFAULT_PRECISION)
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
        BigFloat.new(days.to_f + 2.0, precision: Functions::DEFAULT_PRECISION) # Excel's 1900 date system has a bug treating 1900 as leap year
      when Nil
        nil
      else
        # Fallback: convert to string
        value.to_s
      end
    end
  end
end
