require "compress/zip"
require "xml"

module Sheety
  class ExcelExporter
    # Export sheety internal format to Excel .xlsx file
    #
    # Parameters:
    # - data: Hash(String, Hash(String, Hash(String, Functions::CellValue)))
    #   Format: { "SheetName" => { "A1" => { "value" => ..., "formula" => ... }, ... } }
    # - filename: Output .xlsx file path
    #
    def self.export_to_xlsx(data : Hash(String, Hash(String, Hash(String, Functions::CellValue))), filename : String)
      # Create a temporary directory structure for the xlsx file
      temp_dir = File.join("/tmp", "xlsx_export_#{Random.new.next_int}")
      Dir.mkdir_p(temp_dir)

      # Collect all unique strings for shared string table
      shared_strings = [] of String
      string_to_index = Hash(String, Int32).new

      data.each do |_, sheet_data|
        sheet_data.each do |_, cell_data|
          if cell_data.has_key?("value")
            value = cell_data["value"]
            if value.is_a?(String) && !numeric_string?(value) && !boolean_string?(value)
              unless string_to_index.has_key?(value)
                string_to_index[value] = shared_strings.size
                shared_strings << value
              end
            end
          end
        end
      end

      begin
        # Create directory structure
        xl_dir = File.join(temp_dir, "xl")
        worksheets_dir = File.join(xl_dir, "worksheets")
        rels_dir = File.join(xl_dir, "_rels")
        Dir.mkdir_p(worksheets_dir)
        Dir.mkdir_p(rels_dir)

        # Create [Content_Types].xml
        content_types = String.build do |xml|
          xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
          xml << "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
          xml << "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
          xml << "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
          xml << "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
          data.keys.each_with_index do |_, index|
            xml << "  <Override PartName=\"/xl/worksheets/sheet#{index + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
          end
          xml << "  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n"
          xml << "  <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n"
          xml << "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n"
          xml << "  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n"
          xml << "  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n"
          xml << "</Types>\n"
        end

        # Create workbook.xml
        sheets_xml = data.keys.map_with_index do |sheet_name, index|
          "    <sheet name=\"#{escape_xml(sheet_name)}\" sheetId=\"#{index + 1}\" r:id=\"rId#{index + 1}\"/>"
        end.join("\n")

        workbook = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
        #{sheets_xml}
          </sheets>
        </workbook>
        XML

        # Create sharedStrings.xml
        shared_strings_xml = String.build do |xml|
          xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
          xml << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"#{shared_strings.size}\" uniqueCount=\"#{shared_strings.size}\">\n"
          shared_strings.each do |str|
            xml << "  <si><t>#{escape_xml(str)}</t></si>\n"
          end
          xml << "</sst>\n"
        end

        # Create workbook.xml.rels
        rels = String.build do |xml|
          xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
          xml << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
          data.keys.each_with_index do |_, index|
            xml << "  <Relationship Id=\"rId#{index + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet#{index + 1}.xml\"/>\n"
          end
          # Add shared strings relationship
          xml << "  <Relationship Id=\"rId#{data.keys.size + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\n"
          xml << "</Relationships>\n"
        end

        # Create worksheets
        data.each_with_index do |(sheet_name, sheet_data), index|
          worksheet_xml = generate_worksheet_xml(sheet_data, string_to_index)
          File.write(File.join(worksheets_dir, "sheet#{index + 1}.xml"), worksheet_xml)
        end

        # Create minimal styles.xml
        styles = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="0"/>
          <fonts count="1">
            <font><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
          </fills>
          <borders count="1">
            <border><left/><right/><top/><bottom/></border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        XML

        # Create minimal theme
        theme = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
          <a:themeElements>
            <a:colorScheme name="Office">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
              <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
              <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
              <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
              <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
              <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
              <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
              <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
              <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
              <a:accent6><a:srgbClr val="F79646"/></a:accent6>
            </a:colorScheme>
          </a:themeElements>
        </a:theme>
        XML

        # Create docProps
        doc_props_core = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/">
          <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
          <dc:creator>Sheety</dc:creator>
        </cp:coreProperties>
        XML

        doc_props_app = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
          <Application>Sheety</Application>
        </Properties>
        XML

        # Write all files
        File.write(File.join(temp_dir, "[Content_Types].xml"), content_types)
        File.write(File.join(xl_dir, "workbook.xml"), workbook)
        File.write(File.join(xl_dir, "sharedStrings.xml"), shared_strings_xml)
        File.write(File.join(rels_dir, "workbook.xml.rels"), rels)
        File.write(File.join(xl_dir, "styles.xml"), styles)

        theme_dir = File.join(xl_dir, "theme")
        Dir.mkdir_p(theme_dir)
        File.write(File.join(theme_dir, "theme1.xml"), theme)

        doc_props_dir = File.join(temp_dir, "docProps")
        Dir.mkdir_p(doc_props_dir)
        File.write(File.join(doc_props_dir, "core.xml"), doc_props_core)
        File.write(File.join(doc_props_dir, "app.xml"), doc_props_app)

        # Create .rels file
        dot_rels_dir = File.join(temp_dir, "_rels")
        Dir.mkdir_p(dot_rels_dir)
        dot_rels = <<-XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        XML
        File.write(File.join(dot_rels_dir, ".rels"), dot_rels)

        # Create the zip file
        create_xlsx_zip(temp_dir, filename)
      ensure
        # Clean up temp directory
        FileUtils.rm_rf(temp_dir) if Dir.exists?(temp_dir)
      end
    end

    # Check if a string represents a number
    private def self.numeric_string?(str : String) : Bool
      !!(str =~ /^-?\d+\.?\d*$/)
    end

    # Check if a string represents a boolean
    private def self.boolean_string?(str : String) : Bool
      str == "TRUE" || str == "FALSE"
    end

    # Generate worksheet XML with formula support
    private def self.generate_worksheet_xml(sheet_data : Hash(String, Hash(String, Functions::CellValue)), string_to_index : Hash(String, Int32)) : String
      # Sort cells and organize into rows
      rows_data = Hash(Int32, Array(NamedTuple(cell_ref: String, col: Int32, value: String?, formula: String?, cell_type: String?))).new

      sheet_data.each do |cell_ref, cell_data|
        row_num = parse_row(cell_ref)
        col_num = parse_col(cell_ref)

        rows_data[row_num] ||= [] of NamedTuple(cell_ref: String, col: Int32, value: String?, formula: String?, cell_type: String?)

        value = nil
        formula = nil
        cell_type = nil

        if cell_data.has_key?("formula")
          formula = cell_data["formula"].to_s
          # Remove '=' prefix if present (Excel stores formulas without '=' in the <f> element)
          formula = formula[1..-1] if formula.starts_with?("=")
        elsif cell_data.has_key?("value")
          raw_value = cell_data["value"]
          value_str = value_to_string(raw_value)

          if raw_value.is_a?(Bool)
            cell_type = "b"
            value = value_str
          elsif raw_value.is_a?(Number) || (raw_value.is_a?(String) && numeric_string?(raw_value))
            cell_type = "n"
            value = value_str
          elsif raw_value.is_a?(String) && string_to_index.has_key?(raw_value)
            cell_type = "s"
            value = string_to_index[raw_value].to_s
          else
            # Fallback for other types
            cell_type = nil
            value = value_str
          end
        end

        rows_data[row_num] << {cell_ref: cell_ref, col: col_num, value: value, formula: formula, cell_type: cell_type}
      end

      # Build the XML
      String.build do |xml|
        xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
        xml << "  <sheetData>\n"

        # Sort rows by row number and iterate
        rows_data.to_a.sort_by { |(row_num, _)| row_num }.each do |(_, cells)|
          sorted_cells = cells.sort_by { |cell| cell[:col] }

          row_num = sorted_cells.first[:cell_ref].match(/\d+$/).try(&.[0]) || "1"
          xml << "    <row r=\"#{row_num}\">\n"

          sorted_cells.each do |cell|
            xml << "      <c r=\"#{cell[:cell_ref]}\""

            if formula = cell[:formula]
              # Formula cell - no type attribute
              xml << ">\n"
              xml << "        <f>#{escape_xml(formula)}</f>\n"
              xml << "      </c>\n"
            elsif value = cell[:value]
              if cell_type = cell[:cell_type]
                # Value cell with type
                xml << " t=\"#{cell_type}\">#{format_cell_value(value, cell_type)}</c>\n"
              else
                # Value cell without type (shouldn't happen, but fallback)
                xml << ">#{value}</c>\n"
              end
            else
              xml << "/>\n"
            end
          end

          xml << "    </row>\n"
        end

        xml << "  </sheetData>\n"
        xml << "</worksheet>\n"
      end
    end

    # Format cell value for XML
    private def self.format_cell_value(value : String, cell_type : String) : String
      case cell_type
      when "b"
        "        <v>#{value == "TRUE" ? "1" : "0"}</v>"
      else
        "        <v>#{value}</v>"
      end
    end

    # Create ZIP file from directory structure
    private def self.create_xlsx_zip(source_dir : String, output_file : String)
      File.open(output_file, "w") do |file|
        Compress::Zip::Writer.open(file) do |zip|
          Dir.glob(File.join(source_dir, "**", "*")).sort.each do |filepath|
            next if File.directory?(filepath)

            # Calculate relative path from source_dir
            relative_path = filepath[(source_dir.size + 1)..-1]

            # Add file to zip
            zip.add(relative_path, File.read(filepath))
          end
        end
      end
    end

    # Convert cell value to string representation for Excel
    private def self.value_to_string(value : Functions::CellValue) : String
      case value
      when String
        value
      when Float64
        if value == value.to_i
          value.to_i.to_s
        else
          value.to_s
        end
      when Int32, Int64
        value.to_s
      when Bool
        value.to_s.upcase
      when Nil
        ""
      else
        value.to_s
      end
    end

    # Parse row number from cell reference (e.g., "A1" -> 1)
    private def self.parse_row(cell_ref : String) : Int32
      match = cell_ref.match(/[A-Za-z]+(\d+)/)
      if match
        match[1].to_i
      else
        1
      end
    end

    # Parse column number from cell reference (e.g., "A1" -> 1, "B1" -> 2)
    private def self.parse_col(cell_ref : String) : Int32
      match = cell_ref.match(/([A-Za-z]+)\d+/)
      if match
        col_str = match[1]
        col_num = 0
        col_str.each_char do |char|
          col_num = col_num * 26 + (char.upcase.ord - 'A'.ord + 1)
        end
        col_num
      else
        1
      end
    end

    # Escape XML special characters
    private def self.escape_xml(str : String) : String
      str
        .gsub("&", "&amp;")
        .gsub("<", "&lt;")
        .gsub(">", "&gt;")
        .gsub("\"", "&quot;")
        .gsub("'", "&apos;")
    end
  end
end
