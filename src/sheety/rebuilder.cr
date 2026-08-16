require "big"
require "yaml"
require "uuid"
require "./data_dir"
require "./yaml_parser"
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

    def spreadsheet_uuid=(uuid : String) : self
      @spreadsheet_uuid = uuid
      self
    end

    def intermediate_file=(file : String) : self
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
      file_hash = DataDir.file_hash(source_file)

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
      generator.state_file_path = croupier_state
      generator.kv_store_path = kv_store
      generator.spreadsheet_uuid = spreadsheet_uuid
      generator.original_filename = filename
      initial_values = Hash(String, BigFloat | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(intermediate_file)
      data = YAML.parse(yaml_content)
      process_yaml_data(data, generator, initial_values)

      # Generate Croupier task source code with initial values (non-interactive for rebuild)
      # Use intermediate_file as the source file so the TUI reads from the updated file
      generated = generator.generate_source(initial_values, true, intermediate_file, nil, hash_short)

      if generated.entrypoint.empty?
        STDERR.puts "Error: Failed to generate source code - output is empty"
        return nil
      end

      # Write the entrypoint and any chunk files (large sheets split the task
      # table across files to cap compiler memory)
      CroupierGenerator.write_generated(generated, output_cr)

      # Build the binary
      build_result = Process.run("crystal", ["build", output_cr, "-o", binary_name], output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      unless build_result.success?
        STDERR.puts "\nError: Build failed"
        return nil
      end

      binary_name
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

        # Read the current YAML content and make sure it carries the UUID
        yaml_content = File.read(filename)
        effective_uuid, new_content = YAMLParser.ensure_uuid_in_yaml(yaml_content, uuid)
        File.write(filename, new_content) if new_content != yaml_content
        uuid = effective_uuid
      end

      uuid
    end

    private def process_yaml_data(data : YAML::Any, generator : CroupierGenerator, initial_values : Hash(String, BigFloat | String | Bool))
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
            value = YAMLParser.parse_value(cell_data["value"])
            initial_values[key] = value
          end
        end
      end
    end
  end
end
