require "yaml"
require "./croupier_generator"
require "./importers/excel_importer"
require "./data_dir"
require "openssl"

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

      handle_compile(filename, extra_args)
    end

    private def self.handle_compile(filename : String, extra_args : Array(String))
      unless filename && !filename.empty?
        STDERR.puts "Error: No input file specified"
        STDERR.puts "Usage: sheety <file.(yaml|xlsx)> [--no-interactive]"
        exit 1
      end

      unless File.exists?(filename)
        STDERR.puts "Error: File not found: #{filename}"
        exit 1
      end

      # For .xlsx files, convert to YAML and process as YAML
      ext = File.extname(filename).downcase

      if ext == ".xlsx"
        puts "Converting Excel file to YAML format..."
        yaml_filename = convert_excel_to_yaml(filename)
        # Process the converted YAML file
        handle_yaml_file(yaml_filename, extra_args)
        return
      end

      # Handle YAML files
      handle_yaml_file(filename, extra_args)
    end

    private def self.handle_yaml_file(filename : String, extra_args : Array(String))
      # Check for flags
      compile_only = extra_args.includes?("--compile-only")
      interactive = !extra_args.includes?("--no-interactive")

      # Calculate hash of source file for caching
      file_hash = calculate_file_hash(filename)

      # Use first 16 characters of hash for filename
      hash_short = file_hash[0...16]

      # Output files in data directory tmp
      output_cr = File.join(DataDir.path, "tmp", "#{hash_short}.cr")
      binary_name = File.join(DataDir.path, "tmp", "#{hash_short}")
      croupier_state = File.join(DataDir.path, "tmp", "#{hash_short}.croupier")
      kv_store = File.join(DataDir.path, "tmp", "#{hash_short}.kv")

      # Generate the Crystal source file using CroupierGenerator
      generator = CroupierGenerator.new
      generator.set_state_file_path(croupier_state)
      generator.set_kv_store_path(kv_store)
      initial_values = Hash(String, Float64 | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(filename)
      data = YAML.parse(yaml_content)
      process_yaml_data(data, generator, initial_values)

      # Generate Croupier task source code with initial values
      source_code = generator.generate_source(initial_values, interactive, filename)

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

      if compile_only
        # Print binary path and exit
        puts binary_name
        exit 0
      elsif interactive
        puts "\nLaunching TUI..."
        puts "Press Q to exit\n"

        # Run the binary - it handles its own rebuilding via Process.exec
        run_result = Process.run(binary_name, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

        # Handle non-zero exit codes
        unless run_result.success?
          puts "\nNote: TUI requires a terminal. Run './#{binary_name}' in a terminal to view the spreadsheet."
        end
        exit run_result.exit_code
      else
        puts "\nRun with: #{binary_name}"
      end
    end

    private def self.parse_value(value : YAML::Any) : Functions::CellValue
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

    # Process YAML data and add to generator
    private def self.process_yaml_data(data : YAML::Any, generator : CroupierGenerator, initial_values : Hash(String, Float64 | String | Bool))
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

    # Convert Excel file to YAML format and save to data directory
    private def self.convert_excel_to_yaml(filename : String) : String
      # Get the base filename and change extension to .yaml
      basename = File.basename(filename, ".xlsx")
      yaml_filename = File.join(DataDir.path, "#{basename}.yaml")

      # If file already exists, append a number
      if File.exists?(yaml_filename)
        counter = 1
        loop do
          new_name = File.join(DataDir.path, "#{basename}_#{counter}.yaml")
          break unless File.exists?(new_name)
          counter += 1
        end
        yaml_filename = File.join(DataDir.path, "#{basename}_#{counter}.yaml")
      end

      # Parse Excel file and convert to internal format (Hash)
      workbook = ExcelImporter.parse_xlsx(filename)
      hash_data = ExcelImporter.to_internal_format(workbook)

      # Convert Hash to YAML string manually to ensure proper format
      yaml_string = hash_to_yaml_string(hash_data)

      # Write YAML file
      File.write(yaml_filename, yaml_string)
      puts "Created YAML file: #{yaml_filename}"

      yaml_filename
    end

    # Convert the internal hash format to a YAML string
    private def self.hash_to_yaml_string(data : Hash(String, Hash(String, Hash(String, Functions::CellValue)))) : String
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
                lines << "    #{key}: #{("=".to_s + value).inspect}"
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

    private def self.print_help : Nil
      puts <<-HELP
        Usage: sheety <file.(yaml|xlsx)> [options]

        Compiles a spreadsheet to a standalone binary with interactive TUI.

        Options:
          --compile-only      Build binary and print its path, then exit
          --no-interactive    Generate non-interactive binary (runs once and exits)
          -h, --help          Show this help message

        Input formats:
          .xlsx               Excel 2007+ format (with formula support)
          .yaml, .yml         YAML format

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
