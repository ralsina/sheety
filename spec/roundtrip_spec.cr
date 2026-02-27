require "./spec_helper"
require "../src/sheety"
require "yaml"

describe "Excel Roundtrip" do
  it "preserves data when exporting and importing back" do
    # Create a test YAML file
    test_yaml = <<-YAML
    Sheet1:
      A1:
        value: "10.5"
      A2:
        value: "20"
      A3:
        value: "30"
      A4:
        formula: "=SUM(A1:A3)"
      A5:
        formula: "=A4*2"
      B1:
        value: "Hello"
      B2:
        value: "World"
      C1:
        value: true
      C2:
        value: false
    Sheet2:
      A1:
        value: 100
      A2:
        formula: "=Sheet1!A4"
    _ui_state:
      active_sheet: Sheet1
      active_cell: A1
    YAML

    # Create temp directory for testing
    temp_dir = File.join("/tmp", "sheety_roundtrip_test")
    Dir.mkdir_p(temp_dir) unless Dir.exists?(temp_dir)

    yaml_file = File.join(temp_dir, "roundtrip_test.yaml")
    xlsx_file = File.join(temp_dir, "roundtrip_test.xlsx")
    reimported_yaml = File.join(temp_dir, "roundtrip_test_reimported.yaml")

    begin
      # Write test YAML
      File.write(yaml_file, test_yaml)

      # Load the original data
      original_data = YAML.parse(test_yaml)

      # Export to XLSX using ExcelExporter
      internal_format = Hash(String, Hash(String, Hash(String, Sheety::Functions::CellValue))).new

      original_data.as_h.each do |sheet_name, sheet_data|
        next if sheet_name.as_s == "_ui_state"

        sheet_hash = Hash(String, Hash(String, Sheety::Functions::CellValue)).new
        sheet_data.as_h.each do |cell_ref, cell_data|
          cell_hash = Hash(String, Sheety::Functions::CellValue).new

          cell_data.as_h.each do |key, value|
            case key.as_s
            when "value"
              # Parse value like CLI does
              cell_hash["value"] = parse_roundtrip_value(value)
            when "formula"
              cell_hash["formula"] = value.as_s
            end
          end

          sheet_hash[cell_ref.as_s] = cell_hash
        end

        internal_format[sheet_name.as_s] = sheet_hash
      end

      Sheety::ExcelExporter.export_to_xlsx(internal_format, xlsx_file)

      # Import back using XlsxImporter
      workbook = Sheety::ExcelImporter.parse_xlsx(xlsx_file)
      reimported_internal = Sheety::ExcelImporter.to_internal_format(workbook)

      # Save reimported data to YAML for comparison
      reimported_yaml_content = hash_to_roundtrip_yaml_string(reimported_internal)
      File.write(reimported_yaml, reimported_yaml_content)

      # Load reimported data
      reimported_data = YAML.parse(reimported_yaml_content)

      # Verify each sheet's data matches (compare cell-by-cell)
      original_sheets = original_data.as_h.reject { |k| k.as_s == "_ui_state" }
      reimported_sheets = reimported_data.as_h.reject { |k| k.as_s == "_ui_state" }

      # Check sheet count
      original_sheets.size.should eq reimported_sheets.size

      # Verify each sheet's data matches
      original_sheets.each do |sheet_name, original_sheet|
        reimported_sheet = reimported_sheets[sheet_name]

        original_sheet.as_h.each do |cell_ref, original_cell|
          reimported_cell = reimported_sheet.as_h[cell_ref]

          # Compare values
          original_value = original_cell.as_h["value"]?
          reimported_value = reimported_cell.as_h["value"]?

          if original_value && reimported_value
            # Compare as strings for simplicity
            normalize_roundtrip_value(original_value).should eq(normalize_roundtrip_value(reimported_value))
          end

          # Compare formulas
          original_formula = original_cell.as_h["formula"]?
          reimported_formula = reimported_cell.as_h["formula"]?

          if original_formula && reimported_formula
            # Both should have '=' prefix or both not
            original_formula.to_s.should eq(reimported_formula.to_s)
          end
        end
      end
    ensure
      # Cleanup temp files
      File.delete(yaml_file) if File.exists?(yaml_file)
      File.delete(xlsx_file) if File.exists?(xlsx_file)
      File.delete(reimported_yaml) if File.exists?(reimported_yaml)
      Dir.delete(temp_dir) if Dir.exists?(temp_dir)
    end
  end
end

# Helper to parse value (same as CLI)
private def parse_roundtrip_value(value : YAML::Any) : Sheety::Functions::CellValue
  raw = value.raw

  case raw
  when String
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

# Helper to normalize values for comparison
private def normalize_roundtrip_value(value)
  case value
  when Float64
    if value == value.to_i
      value.to_i
    else
      value
    end
  when String
    value
  when Bool
    value.to_s.upcase
  when Nil
    ""
  else
    value.to_s
  end
end

# Helper to convert internal format to YAML string (from cli.cr)
private def hash_to_roundtrip_yaml_string(data : Hash(String, Hash(String, Hash(String, Sheety::Functions::CellValue)))) : String
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

  lines.join("\n")
end
