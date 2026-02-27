require "yaml"
require "uuid"
require "./importers/excel_importer"
require "./importers/excel_exporter"
require "./data_dir"
require "./croupier_generator"

module Sheety
  # Spreadsheet data structures and converters
  module Spreadsheet
    # Cell data structure
    alias CellValue = String | Float64 | Bool | Nil
    alias CellData = Hash(String, CellValue)     # "formula" => "...", "value" => ...
    alias SheetData = Hash(String, CellData)     # "A1" => {...}, "B1" => {...}
    alias WorkbookData = Hash(String, SheetData) # "Sheet1" => {...}, "Sheet2" => {...}

    # Result structure for read operations
    class SpreadsheetFile
      property data : WorkbookData
      property uuid : String
      property source_file : String

      def initialize(@data : WorkbookData, @uuid : String, @source_file : String)
      end
    end

    # Read a file and return the data with metadata
    def self.read_with_metadata(file_path : String) : SpreadsheetFile
      ext = File.extname(file_path).downcase

      # Read the data
      data = case ext
             when ".yaml", ".yml"
               read_yaml(file_path)
             when ".xlsx"
               read_excel(file_path)
             else
               raise "Unsupported input format: #{ext}"
             end

      # Get or create UUID
      uuid = get_or_create_uuid_for_data(data, file_path)

      SpreadsheetFile.new(data, uuid, file_path)
    end

    # Read a file and return the data
    def self.read(file_path : String) : WorkbookData
      read_with_metadata(file_path).data
    end

    # Create an empty spreadsheet file
    def self.create_empty(file_path : String) : Nil
      ext = File.extname(file_path).downcase

      case ext
      when ".yaml", ".yml"
        create_empty_yaml(file_path)
      when ".xlsx"
        create_empty_excel(file_path)
      else
        # Default to YAML for unknown extensions
        create_empty_yaml(file_path)
      end
    end

    # Get or create UUID for spreadsheet data
    private def self.get_or_create_uuid_for_data(data : WorkbookData, file_path : String) : String
      uuid = nil

      # Check if data already has _ui_state with uuid (from read_yaml)
      ui_state = data["_ui_state"]?
      if ui_state
        uuid_value = ui_state["spreadsheet_uuid"]?
        if uuid_value
          uuid = uuid_value.to_s
        end
      end

      # Create new UUID if needed
      unless uuid
        uuid = UUID.random.to_s

        # If it's a YAML file, write the UUID back to it
        if File.extname(file_path).downcase == ".yaml"
          # Add UUID to _ui_state section
          yaml_content = File.read(file_path)

          if yaml_content.includes?("_ui_state:")
            if yaml_content.includes?("spreadsheet_uuid:")
              # Already has it, try to extract
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
              File.write(file_path, yaml_content)
            end
          else
            # Add _ui_state section at the end
            yaml_content = yaml_content + "\n_ui_state:\n  spreadsheet_uuid: #{uuid}\n"
            File.write(file_path, yaml_content)
          end
        end
      end

      uuid
    end

    # Write the in-memory representation to a file
    def self.write(data : WorkbookData, file_path : String, source_file : String? = nil) : Nil
      ext = File.extname(file_path).downcase

      case ext
      when ".yaml", ".yml"
        write_yaml(data, file_path)
      when ".xlsx"
        write_excel(data, file_path)
      when ".cr"
        write_crystal_source(data, file_path, source_file)
      when ".sheety"
        write_binary(data, file_path, source_file)
      else
        raise "Unsupported output format: #{ext}"
      end
    end

    # Convert between any two formats
    def self.convert(from_file : String, to_file : String) : Nil
      data = read(from_file)
      write(data, to_file, from_file)
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

    # Write Crystal source code format
    private def self.write_crystal_source(data : WorkbookData, file_path : String, source_file : String?) : Nil
      generator = CroupierGenerator.new
      initial_values = Hash(String, Float64 | String | Bool).new

      # Convert data to Croupier format
      data.each do |sheet_name, sheet_data|
        sheet_data.each do |cell_ref, cell_data|
          key = "#{sheet_name}!#{cell_ref}"

          if cell_data.has_key?("formula")
            formula = cell_data["formula"].to_s
            generator.add_formula(cell_ref.to_s, formula, sheet_name.to_s)
          elsif cell_data.has_key?("value")
            value = cell_data["value"]
            # Skip nil values
            next if value.nil?
            # Convert to appropriate type
            case value
            when String, Float64, Bool
              initial_values[key] = value
            else
              initial_values[key] = value.to_s
            end
          end
        end
      end

      # Generate source code
      source_code = generator.generate_source(initial_values, true, source_file || file_path, nil)

      if source_code.empty?
        raise "Failed to generate source code"
      end

      File.write(file_path, source_code)

      # Format the generated Crystal file
      format_result = Process.run("crystal", ["tool", "format", file_path],
        output: Process::Redirect::Inherit,
        error: Process::Redirect::Inherit)

      unless format_result.success?
        STDERR.puts "Warning: Failed to format generated Crystal file"
      end
    end

    # Write compiled binary format
    private def self.write_binary(data : WorkbookData, file_path : String, source_file : String?) : Nil
      generator = CroupierGenerator.new
      initial_values = Hash(String, Float64 | String | Bool).new

      # Convert data to Croupier format
      data.each do |sheet_name, sheet_data|
        sheet_data.each do |cell_ref, cell_data|
          key = "#{sheet_name}!#{cell_ref}"

          if cell_data.has_key?("formula")
            formula = cell_data["formula"].to_s
            generator.add_formula(cell_ref.to_s, formula, sheet_name.to_s)
          elsif cell_data.has_key?("value")
            value = cell_data["value"]
            # Skip nil values
            next if value.nil?
            # Convert to appropriate type
            case value
            when String, Float64, Bool
              initial_values[key] = value
            else
              initial_values[key] = value.to_s
            end
          end
        end
      end

      # Generate source code
      source_file_for_binary = source_file || file_path
      source_code = generator.generate_source(initial_values, true, source_file_for_binary, nil)

      if source_code.empty?
        raise "Failed to generate source code"
      end

      # Compile
      temp_source = File.join(DataDir.path, "tmp", "#{File.basename(file_path, File.extname(file_path))}.cr")
      File.write(temp_source, source_code)

      compile_result = Process.run("crystal", ["build", "-Dpreview_mt", "-Dno_embedded_files", temp_source, "-o", file_path],
        output: Process::Redirect::Inherit,
        error: Process::Redirect::Inherit)

      unless compile_result.success?
        raise "Compilation failed"
      end
    end

    # Create an empty YAML spreadsheet
    private def self.create_empty_yaml(file_path : String) : Nil
      content = <<-YAML
        Sheet1:
          A1:
            value: 0
        _ui_state:
          spreadsheet_uuid: #{UUID.random.to_s}
        YAML

      File.write(file_path, content.strip)
    end

    # Create an empty Excel spreadsheet
    private def self.create_empty_excel(file_path : String) : Nil
      # Create via YAML intermediate then convert
      temp_yaml = File.join(DataDir.path, "tmp", "empty_#{UUID.random.to_s}.yaml")
      create_empty_yaml(temp_yaml)
      convert(temp_yaml, file_path)
      File.delete(temp_yaml)
    end
  end
end
