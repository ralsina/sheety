require "yaml"
require "openssl"
require "uuid"
require "./data_dir"
# Require only the specific modules we need, not the main sheety.cr which runs the CLI
require "./ast"
require "./ast_builder"
require "./parser"
require "./tokens/operand"
require "./tokens/operator"
require "./tokens/parenthesis"
require "./functions/registry"
require "./code_generator"
require "./dependency_extractor"
require "./croupier_generator"
require "./errors"

module Sheety
  # Handles rebuilding the binary when formulas are edited
  # This is used by the generated TUI to rebuild in-process instead of spawning a subprocess
  class Rebuilder
    @original_filename : String
    @spreadsheet_uuid : String?
    @intermediate_file : String?

    def initialize(@original_filename : String)
    end

    def set_spreadsheet_uuid(uuid : String) : self
      @spreadsheet_uuid = uuid
      self
    end

    def set_intermediate_file(file : String) : self
      @intermediate_file = file
      self
    end

    # Rebuild and return the path to the new binary
    def rebuild : String?
      filename = @original_filename

      unless File.exists?(filename)
        STDERR.puts "Error: File not found: #{filename}"
        return nil
      end

      # Ensure data directory exists and has required files
      DataDir.ensure
      DataDir.ensure_shard_yml
      DataDir.ensure_dependencies
      DataDir.extract_embedded_files

      # Determine the source file to use for reading data
      # If we have an intermediate file, use that; otherwise use the original filename
      source_file = @intermediate_file || filename

      # Get or create persistent UUID for this spreadsheet
      # Only try to read from source_file if we don't already have a UUID
      spreadsheet_uuid = @spreadsheet_uuid || get_or_create_spreadsheet_uuid(source_file)

      # Calculate hash of source file for caching (for binary naming)
      # Use the intermediate file if available, otherwise the original file
      file_hash = calculate_file_hash(source_file)

      # Use first 16 characters of hash for binary/source filenames
      hash_short = file_hash[0...16]

      # Output files in data directory tmp (binary uses hash)
      output_cr = File.join(DataDir.path, "tmp", "#{hash_short}.cr")
      binary_name = File.join(DataDir.path, "tmp", "#{hash_short}")

      # State files use persistent UUID (survive across rebuilds)
      croupier_state = File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.croupier")
      kv_store = File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.kv")

      # Intermediate save file for auto-saves (uses UUID to avoid conflicts)
      intermediate_file = @intermediate_file || File.join(DataDir.path, "#{spreadsheet_uuid}.yaml")

      # Copy original file to intermediate file if it doesn't exist or is outdated
      # Only do this if we don't already have an intermediate file set
      if !@intermediate_file && (!File.exists?(intermediate_file) || File.info(filename).modification_time > File.info(intermediate_file).modification_time)
        FileUtils.cp(filename, intermediate_file)
      end

      # Generate the Crystal source file using CroupierGenerator
      generator = CroupierGenerator.new
      generator.set_state_file_path(croupier_state)
      generator.set_kv_store_path(kv_store)
      generator.set_spreadsheet_uuid(spreadsheet_uuid)
      generator.set_original_filename(filename)
      initial_values = Hash(String, Float64 | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(intermediate_file)
      data = YAML.parse(yaml_content)
      process_yaml_data(data, generator, initial_values)

      # Generate Croupier task source code with initial values (non-interactive for rebuild)
      source_code = generator.generate_source(initial_values, true, filename, intermediate_file)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code - output is empty"
        return nil
      end

      # Write the source file
      File.write(output_cr, source_code)

      # Build the binary
      build_result = Process.run("crystal", ["build", "-Dpreview_mt", output_cr, "-o", binary_name], output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      unless build_result.success?
        STDERR.puts "\nError: Build failed"
        return nil
      end

      binary_name
    end

    private def calculate_file_hash(filename : String) : String
      digest = OpenSSL::Digest.new("SHA256")
      File.open(filename, "rb") do |file|
        digest.update(file)
      end
      digest.final.hexstring
    end

    private def get_or_create_spreadsheet_uuid(filename : String) : String
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
            # Already has spreadsheet_uuid, try to extract it
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

    private def process_yaml_data(data : YAML::Any, generator : CroupierGenerator, initial_values : Hash(String, Float64 | String | Bool))
      data.as_h.each do |sheet_name, sheet_data|
        # Skip UI metadata
        next if sheet_name.as_s == "_ui_state"

        sheet_data.as_h.each do |cell_ref, cell_data|
          cell_data = cell_data.as_h
          key = "#{sheet_name}!#{cell_ref}"

          if cell_data.has_key?("formula")
            formula = cell_data["formula"].to_s
            generator.add_formula(cell_ref.to_s, formula, sheet_name.to_s)
          elsif cell_data.has_key?("value")
            value = parse_value(cell_data["value"])
            initial_values[key] = value
          end
        end
      end
    end

    private def parse_value(value : YAML::Any) : Functions::CellValue
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
