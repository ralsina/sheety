require "yaml"
require "uuid"
require "./importers/excel_importer"
require "./importers/excel_exporter"
require "./data_dir"

module Sheety
  # Spreadsheet data structures and converters
  module Spreadsheet
    # Cell data structure
    alias CellValue = String | Float64 | Bool | Nil
    alias CellData = Hash(String, CellValue)     # "formula" => "...", "value" => ...
    alias SheetData = Hash(String, CellData)     # "A1" => {...}, "B1" => {...}
    alias WorkbookData = Hash(String, SheetData) # "Sheet1" => {...}, "Sheet2" => {...}

    # Get or create spreadsheet UUID from a YAML file
    def self.get_or_create_spreadsheet_uuid(filename : String) : String
      uuid = nil

      # Try to read existing UUID from YAML
      begin
        yaml_content = File.read(filename)
        data = YAML.parse(yaml_content)

        if data.as_h? && data["_ui_state"]? && data["_ui_state"]["spreadsheet_uuid"]?
          uuid = data["_ui_state"]["spreadsheet_uuid"].as_s
        end
      rescue
        # If parsing fails, we'll create a new UUID
      end

      # If no UUID exists, create one and add it to the YAML
      unless uuid
        uuid = UUID.random.to_s

        # Read the current YAML content
        yaml_content = File.read(filename)

        # Check if _ui_state section exists and if it has spreadsheet_uuid
        if yaml_content.includes?("_ui_state:")
          if yaml_content.includes?("spreadsheet_uuid:")
            # Already has spreadsheet_uuid, don't add it again
            # Try to extract it
            yaml_content.each_line do |line|
              if line.includes?("spreadsheet_uuid:")
                match = line.match(/spreadsheet_uuid:\s*(\S+)/)
                if match
                  uuid = match[1]
                  break
                end
              end
            end
          else
            # Append the UUID to existing _ui_state section
            yaml_content = yaml_content.gsub(/(_ui_state:)/, "\\1\n  spreadsheet_uuid: #{uuid}")
            File.write(filename, yaml_content)
          end
        else
          # Add _ui_state section at the end
          yaml_content = yaml_content + "\n_ui_state:\n  spreadsheet_uuid: #{uuid}\n"
          File.write(filename, yaml_content)
        end
      end

      uuid
    end

    # Read a file and return the in-memory representation
    def self.read(file_path : String) : WorkbookData
      ext = File.extname(file_path).downcase

      case ext
      when ".yaml", ".yml"
        read_yaml(file_path)
      when ".xlsx"
        read_excel(file_path)
      else
        raise "Unsupported input format: #{ext}"
      end
    end

    # Write the in-memory representation to a file
    def self.write(data : WorkbookData, file_path : String) : Nil
      ext = File.extname(file_path).downcase

      case ext
      when ".yaml", ".yml"
        write_yaml(data, file_path)
      when ".xlsx"
        write_excel(data, file_path)
      else
        raise "Unsupported output format: #{ext}"
      end
    end

    # Convert between any two formats
    def self.convert(from_file : String, to_file : String) : Nil
      data = read(from_file)
      write(data, to_file)
      puts "Converted: #{from_file} -> #{to_file}"
    end

    # Read YAML format
    private def self.read_yaml(file_path : String) : WorkbookData
      yaml_content = File.read(file_path)
      parsed_data = YAML.parse(yaml_content)

      workbook = WorkbookData.new

      parsed_data.as_h.each do |sheet_name, sheet_data|
        next if sheet_name.as_s == "_ui_state"

        sheet = SheetData.new
        sheet_data.as_h.each do |cell_ref, cell_data|
          cell = CellData.new

          cell_data.as_h.each do |key, value|
            case key.as_s
            when "formula"
              cell["formula"] = value.as_s
            when "value"
              cell["value"] = parse_yaml_value(value)
            end
          end

          sheet[cell_ref.as_s] = cell
        end

        workbook[sheet_name.as_s] = sheet
      end

      workbook
    end

    # Write YAML format
    private def self.write_yaml(data : WorkbookData, file_path : String) : Nil
      lines = [] of String

      data.each do |sheet_name, sheet_data|
        lines << "#{sheet_name}:"
        sheet_data.each do |cell_ref, cell_data|
          lines << "  #{cell_ref}:"
          cell_data.each do |key, value|
            case value
            when String
              # Add = prefix to formulas if missing
              if key == "formula" && !value.starts_with?("=")
                lines << "    #{key}: #{("=" + value).inspect}"
              else
                lines << "    #{key}: #{value.inspect}"
              end
            when Float64
              if value == value.to_i
                lines << "    #{key}: #{value.to_i}"
              else
                lines << "    #{key}: #{value}"
              end
            when Bool
              lines << "    #{key}: #{value}"
            when Nil
              # Skip nil values
            else
              lines << "    #{key}: #{value.inspect}"
            end
          end
        end
      end

      File.open(file_path, "w") do |file|
        file.print(lines.join("\n"))
        file.flush
        file.fsync
      end
    end

    # Read Excel format
    private def self.read_excel(file_path : String) : WorkbookData
      workbook = ExcelImporter.parse_xlsx(file_path)
      internal_format = ExcelImporter.to_internal_format(workbook)

      # Convert internal format to our WorkbookData structure
      workbook_data = WorkbookData.new

      internal_format.each do |sheet_name, sheet_data|
        sheet = SheetData.new
        sheet_data.each do |cell_ref, cell_data|
          cell = CellData.new
          cell_data.each do |key, value|
            # Convert Functions::CellValue to our CellValue type
            cell[key] = convert_cell_value(value)
          end
          sheet[cell_ref] = cell
        end
        workbook_data[sheet_name] = sheet
      end

      workbook_data
    end

    # Convert Functions::CellValue to our CellValue type
    private def self.convert_cell_value(value : Sheety::Functions::CellValue) : CellValue
      case value
      when Sheety::Functions::ErrorValue
        value.to_s
      when Nil
        nil
      when String, Float64, Bool
        value
      else
        value.to_s
      end
    end

    # Write Excel format
    private def self.write_excel(data : WorkbookData, file_path : String) : Nil
      # Convert to internal format expected by ExcelExporter
      internal_format = Hash(String, Hash(String, Hash(String, Functions::CellValue))).new

      data.each do |sheet_name, sheet_data|
        sheet_hash = Hash(String, Hash(String, Functions::CellValue)).new
        sheet_data.each do |cell_ref, cell_data|
          cell_hash = Hash(String, Functions::CellValue).new
          cell_data.each do |key, value|
            cell_hash[key] = value
          end
          sheet_hash[cell_ref] = cell_hash
        end
        internal_format[sheet_name] = sheet_hash
      end

      puts "Exporting to Excel: #{file_path}"
      ExcelExporter.export_to_xlsx(internal_format, file_path)
      puts "Export complete!"
    end

    # Parse YAML value to appropriate type
    private def self.parse_yaml_value(value : YAML::Any) : CellValue
      raw = value.raw

      case raw
      when String
        # Check if it's a boolean
        if raw == "true"
          true
        elsif raw == "false"
          false
        else
          raw
        end
      when Int32, Int64
        raw.to_f
      when Float64
        raw
      when Bool
        raw
      else
        raw.to_s
      end
    end
  end
end
