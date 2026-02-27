require "yaml"
require "./croupier_generator"
require "./spreadsheet"
require "./yaml_parser"
require "./data_dir"
require "uuid"

module Sheety
  # Handles code generation and compilation for spreadsheets
  class Builder
    # Convert between formats (YAML, Excel, Crystal source, compiled binary)
    def self.convert(input_file : String, output_file : String) : Nil
      ext = File.extname(output_file).downcase

      case ext
      when ".xlsx", ".yaml", ".yml"
        # Use Spreadsheet module for data formats
        Spreadsheet.convert(input_file, output_file)
      when ".cr"
        # Generate Crystal source code
        generate_crystal_source(input_file, output_file)
      when ".sheety"
        # Generate and compile binary
        compile_to_binary(input_file, output_file)
      else
        STDERR.puts "Error: Unsupported output format: #{ext}"
        STDERR.puts "Supported formats: .xlsx, .yaml, .yml, .cr, .sheety"
        exit 1
      end

      puts "Converted: #{input_file} -> #{output_file}"
    end

    # Generate Crystal source code from YAML
    private def self.generate_crystal_source(input_file : String, output_file : String) : Nil
      generator = CroupierGenerator.new
      initial_values = Hash(String, Float64 | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(input_file)
      data = YAML.parse(yaml_content)
      YAMLParser.process_yaml_data(data, generator, initial_values)

      # Generate source code
      source_code = generator.generate_source(initial_values, true, input_file, nil)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code"
        exit 1
      end

      File.write(output_file, source_code)
    end

    # Compile to binary
    private def self.compile_to_binary(input_file : String, output_file : String) : Nil
      generator = CroupierGenerator.new
      initial_values = Hash(String, Float64 | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(input_file)
      data = YAML.parse(yaml_content)
      YAMLParser.process_yaml_data(data, generator, initial_values)

      # Generate source code
      # For .sheety output, use the output binary path as source_file
      # so the binary knows its own name for save operations
      source_file_for_binary = File.extname(output_file) == ".sheety" ? output_file : input_file

      # For .sheety binaries, also set up an intermediate file for formula edits
      # We need a UUID for the intermediate file - generate one based on output filename
      intermediate_file_for_binary = nil
      if File.extname(output_file) == ".sheety"
        # Use a UUID-based intermediate file for .sheety binaries
        # This allows formula edits to be saved without conflicts
        binary_uuid = UUID.random.to_s
        intermediate_file_for_binary = File.join(DataDir.path, "#{binary_uuid}.yaml")

        # Also write the initial YAML data to the intermediate file
        FileUtils.cp(input_file, intermediate_file_for_binary)

        # Set the UUID in the generator so it's consistent across rebuilds
        generator.set_spreadsheet_uuid(binary_uuid)
      end

      source_code = generator.generate_source(initial_values, true, source_file_for_binary, intermediate_file_for_binary)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code"
        exit 1
      end

      # Compile
      temp_source = File.join(DataDir.path, "tmp", "#{File.basename(output_file, File.extname(output_file))}.cr")
      File.write(temp_source, source_code)

      compile_args = ["build", "-Dpreview_mt", "-Dno_embedded_files", temp_source, "-o", output_file]

      compile_result = Process.run("crystal", compile_args,
        output: Process::Redirect::Inherit,
        error: Process::Redirect::Inherit)

      unless compile_result.success?
        STDERR.puts "Error: Compilation failed"
        exit 1
      end
    end
  end
end
