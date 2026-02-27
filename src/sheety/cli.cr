require "yaml"
require "./croupier_generator"
require "./spreadsheet"
require "./data_dir"
require "openssl"
require "uuid"
require "./importers/excel_importer"

module Sheety
  class CLI
    # Calculate SHA256 hash of a file for content-based caching
    private def self.calculate_file_hash(filename : String) : String
      digest = OpenSSL::Digest.new("SHA256")
      File.open(filename, "rb") do |file|
        digest.update(file)
      end
      digest.final.hexstring
    end

    # Get or create a persistent UUID for the spreadsheet
    # The UUID is stored in the YAML file's _ui_state section
    # and persists across rebuilds as long as the file exists
    def self.run(args : Array(String))
      # Ensure data directory exists on startup
      DataDir.ensure
      DataDir.ensure_shard_yml
      DataDir.ensure_dependencies
      DataDir.extract_embedded_files

      if args.size == 0 || args.includes?("-h") || args.includes?("--help")
        print_help
        exit 0
      end

      filename = args[0]
      extra_args = args[1..]? || Array(String).new

      handle_file(filename, extra_args)
    end

    private def self.handle_file(filename : String, extra_args : Array(String))
      if !filename || filename.empty?
        STDERR.puts "Error: No input file specified"
        STDERR.puts "Usage: sheety <file.(yaml|xlsx)> [options]"
        exit 1
      end

      # If file doesn't exist, create an empty spreadsheet
      unless File.exists?(filename)
        puts "Creating new spreadsheet: #{filename}"
        Spreadsheet.create_empty(filename)
      end

      # Check for flags
      save_to = extra_args.find(&.starts_with?("--save-to="))
      save_to = save_to.try(&.split('=').last) if save_to

      # Handle --save-to flag for direct format conversion
      if save_to
        Spreadsheet.convert(filename, save_to)
        return
      end

      # Interactive mode: build binary and launch TUI
      # Read the spreadsheet data with metadata (works with any format)
      spreadsheet_file = Spreadsheet.read_with_metadata(filename)
      data = spreadsheet_file.data
      spreadsheet_uuid = spreadsheet_file.uuid

      # For non-YAML files, we need a YAML intermediate file
      if File.extname(spreadsheet_file.source_file).downcase != ".yaml"
        temp_yaml = File.join(DataDir.path, "tmp", "#{UUID.random.to_s}.yaml")
        Spreadsheet.write(data, temp_yaml)
        filename = temp_yaml
      else
        filename = spreadsheet_file.source_file
      end

      # Calculate hash of source file for caching (for binary naming)
      file_hash = calculate_file_hash(filename)

      # Use first 16 characters of hash for binary/source filenames
      hash_short = file_hash[0...16]

      # Output files in data directory tmp (binary uses hash)
      output_cr = File.join(DataDir.path, "tmp", "#{hash_short}.cr")
      binary_name = File.join(DataDir.path, "tmp", "#{hash_short}")

      # State files use persistent UUID (survive across rebuilds)
      croupier_state = File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.croupier")
      kv_store = File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.kv")

      # Intermediate save file for auto-saves (uses UUID to avoid conflicts)
      intermediate_file = File.join(DataDir.path, "#{spreadsheet_uuid}.yaml")

      # Copy source file to intermediate file if it doesn't exist
      if !File.exists?(intermediate_file) || File.info(filename).modification_time > File.info(intermediate_file).modification_time
        FileUtils.cp(filename, intermediate_file)
      end

      # Generate the Crystal source file using CroupierGenerator
      generator = CroupierGenerator.new
      generator.set_state_file_path(croupier_state)
      generator.set_kv_store_path(kv_store)
      generator.set_spreadsheet_uuid(spreadsheet_uuid)
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

      # Generate Croupier task source code with initial values
      source_code = generator.generate_source(initial_values, true, filename, intermediate_file)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code - output is empty"
        exit 1
      end

      # Check if source file already exists
      if File.exists?(output_cr)
        puts "Using cached source: #{output_cr}"
      else
        # Write the source file
        File.write(output_cr, source_code)
      end

      # Check if binary already exists and is newer than source
      if File.exists?(binary_name) && File.info(binary_name).modification_time >= File.info(output_cr).modification_time
        puts "Using cached binary: #{binary_name}"
        build_result = Process::Status.new(0) # Simulate success
      else
        # Build the binary
        puts "Building #{binary_name}..."
        build_result = Process.run("crystal", ["build", "-Dpreview_mt", output_cr, "-o", binary_name], output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)
      end

      unless build_result.success?
        STDERR.puts "\nError: Build failed"
        exit 1
      end

      puts "Built successfully: #{binary_name}"

      puts "\nLaunching TUI..."
      puts "Press Q to exit\n"

      # Run the binary - it handles its own rebuilding via Process.exec
      run_result = Process.run(binary_name, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      # Handle non-zero exit codes
      unless run_result.success?
        puts "\nNote: TUI requires a terminal. Run './#{binary_name}' in a terminal to view the spreadsheet."
      end
      exit run_result.exit_code
    end

    private def self.print_help : Nil
      puts <<-HELP
        Usage: sheety <file.(yaml|xlsx)> [options]

        Compiles a spreadsheet to a standalone binary with interactive TUI.

        Options:
          --save-to=FILE      Convert and save to specified format (extension determines type)
          -h, --help          Show this help message

        Output formats (via --save-to):
          .xlsx               Excel file
          .yaml, .yml         YAML file
          .cr                 Crystal source code
          .sheety             Interactive binary with TUI

        Input formats:
          .xlsx               Excel 2007+ format (with formula support)
          .yaml, .yml         YAML format

        Examples:
          sheety data.yaml --save-to=data.xlsx              # Convert to Excel
          sheety data.xlsx --save-to=data.yaml              # Convert to YAML
          sheety data.yaml --save-to=data.sheety            # Compile to binary
          sheety data.yaml --save-to=source.cr              # Generate source code

        YAML Format:
          SheetName:
            A1:
              value: 10
            A2:
              formula: "=SUM(A1:A3)"
      HELP
    end
  end
end
