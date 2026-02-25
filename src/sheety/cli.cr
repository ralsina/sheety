require "yaml"
require "./croupier_generator"
require "./importers/excel_importer"

module Sheety
  class CLI
    def self.run(args : Array(String))
      if args.size == 0 || args.includes?("-h") || args.includes?("--help")
        print_help
        exit 0
      end

      command = args[0]
      filename = args[1]?

      case command
      when "compile"
        extra_args = args[2..]? || Array(String).new
        handle_compile(filename || "", extra_args)
      else
        # Default: evaluate and print
        handle_evaluate(command, filename)
      end
    end

    private def self.handle_compile(filename : String, extra_args : Array(String))
      unless filename
        STDERR.puts "Error: No input file specified"
        STDERR.puts "Usage: sheety compile <input.yaml> [output.cr] [--no-interactive]"
        exit 1
      end

      unless File.exists?(filename)
        STDERR.puts "Error: File not found: #{filename}"
        exit 1
      end

      # Check for --no-interactive flag to disable interactive mode
      interactive = !extra_args.includes?("--no-interactive")

      # For .xlsx files, convert to YAML and save in examples/ first
      ext = File.extname(filename).downcase
      actual_filename = filename

      if ext == ".xlsx"
        puts "Converting Excel file to YAML format..."
        actual_filename = convert_excel_to_yaml(filename)
      end

      # Generate output filename if not specified
      output_base = extra_args[0]? || actual_filename.gsub(/\.yaml$/, "").gsub(/\.yml$/, "")

      # Ensure .cr extension for source file
      output_cr = if output_base.ends_with?(".cr")
                    output_base
                  else
                    output_base + ".cr"
                  end

      # Generate the Crystal source file using CroupierGenerator
      generator = CroupierGenerator.new
      initial_values = Hash(String, Float64 | String | Bool).new

      # Load YAML file and process
      yaml_content = File.read(actual_filename)
      data = YAML.parse(yaml_content)
      process_yaml_data(data, generator, initial_values)

      # Generate Croupier task source code with initial values
      source_code = generator.generate_source(initial_values, interactive, actual_filename)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code - output is empty"
        exit 1
      end

      # Write the source file
      File.write(output_cr, source_code)

      # Determine binary name (remove .cr extension but keep directory)
      binary_name = output_cr.gsub(/\.cr$/, "")

      # Build the binary
      puts "Building #{binary_name}..."
      build_result = Process.run("crystal", ["build", output_cr, "-o", binary_name], output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      unless build_result.success?
        STDERR.puts "\nError: Build failed"
        exit 1
      end

      puts "Built successfully: #{binary_name}"

      if interactive
        puts "\nLaunching TUI..."
        puts "Press Q to exit\n"

        # Loop to handle recompiles when formulas are edited
        loop do
          # Run the binary
          run_result = Process.run(binary_name, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

          # Exit code 42 means formula was edited, need to recompile
          if run_result.exit_code == 42
            puts "\nFormula edited, regenerating and recompiling...\n"

            # Add a small delay to ensure YAML file is fully written
            sleep(0.2.seconds)

            # Delete the old binary and generated source to ensure clean rebuild
            File.delete(binary_name) if File.exists?(binary_name)
            File.delete(output_cr) if File.exists?(output_cr)

            # Regenerate the Crystal source from the updated YAML
            # We need to recreate the generator and regenerate the source code
            data = YAML.parse(File.read(filename))
            new_initial_values = Hash(String, Float64 | String | Bool).new

            data.as_h.each do |sheet_name, sheet_data|
              # Skip UI metadata
              next if sheet_name.as_s == "_ui_state"

              sheet_data.as_h.each do |cell_ref, cell_data|
                cell_data = cell_data.as_h
                key = "#{sheet_name}!#{cell_ref}"

                if cell_data.has_key?("formula")
                  # Formulas will be regenerated
                elsif cell_data.has_key?("value")
                  value = parse_value(cell_data["value"])
                  new_initial_values[key] = value
                end
              end
            end

            # Create new generator and regenerate source code
            new_generator = CroupierGenerator.new
            data.as_h.each do |sheet_name, sheet_data|
              # Skip UI metadata
              next if sheet_name.as_s == "_ui_state"

              sheet_data.as_h.each do |cell_ref, cell_data|
                cell_data = cell_data.as_h
                if cell_data.has_key?("formula")
                  new_generator.add_formula(cell_ref.to_s, cell_data["formula"].to_s, sheet_name.to_s)
                end
              end
            end

            new_source_code = new_generator.generate_source(new_initial_values, true, filename)

            if new_source_code.empty?
              STDERR.puts "Error: Failed to regenerate source code"
              exit 1
            end

            # Write the regenerated source
            File.write(output_cr, new_source_code)

            # Now build the binary
            build_result = Process.run("crystal", ["build", output_cr, "-o", binary_name], output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

            unless build_result.success?
              STDERR.puts "\nError: Recompile failed"
              exit 1
            end

            puts "Recompiled successfully, restarting...\n"
            # Continue loop to restart TUI
          else
            # Normal exit or error
            unless run_result.success?
              puts "\nNote: TUI requires a terminal. Run './#{binary_name}' in a terminal to view the spreadsheet."
            end
            exit run_result.exit_code
          end
        end
      else
        puts "\nRun with: ./#{binary_name}"
      end
    end

    private def self.handle_evaluate(command : String, filename : String?)
      if filename.nil?
        # Running default evaluation mode without explicit command
        filename = command
      end

      unless File.exists?(filename)
        STDERR.puts "Error: File not found: #{filename}"
        exit 1
      end

      # Load and process the file
      evaluator = load_file(filename)

      # Calculate all formulas
      evaluator.calculate_all

      # Print results
      print_results(evaluator)

      exit 0
    end

    private def self.load_file(filename : String) : Evaluator
      ext = File.extname(filename).downcase

      case ext
      when ".xlsx"
        load_excel_file(filename)
      when ".yaml", ".yml"
        load_yaml_file(filename)
      else
        STDERR.puts "Error: Unsupported file format: #{ext}"
        STDERR.puts "Supported formats: .xlsx, .yaml, .yml"
        exit 1
      end
    end

    private def self.load_excel_file(filename : String) : Evaluator
      workbook = ExcelImporter.parse_xlsx(filename)
      evaluator = Evaluator.new

      workbook.sheets.each do |sheet|
        sheet.cells.each do |cell|
          if formula = cell.formula
            evaluator.set_formula(cell.reference, formula, sheet.name)
          elsif value = cell.value
            evaluator.set(cell.reference, value, sheet.name)
          end
        end
      end

      evaluator
    end

    private def self.load_yaml_file(filename : String) : Evaluator
      data = YAML.parse(File.read(filename))
      load_yaml_data_to_evaluator(data)
    end

    private def self.load_yaml_data_to_evaluator(data : YAML::Any) : Evaluator
      evaluator = Evaluator.new

      data.as_h.each do |sheet_name, sheet_data|
        sheet_data.as_h.each do |cell_ref, cell_data|
          cell_data = cell_data.as_h

          if cell_data.has_key?("value")
            value = parse_value(cell_data["value"])
            evaluator.set(cell_ref.to_s, value, sheet_name.to_s)
          elsif cell_data.has_key?("formula")
            evaluator.set_formula(cell_ref.to_s, cell_data["formula"].to_s, sheet_name.to_s)
          end
        end
      end

      evaluator
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

    # Parse value from CellValue type (used for Excel import)
    private def self.parse_cell_value(value : Functions::CellValue) : Functions::CellValue
      value
    end

    # Process Excel file data and add to generator
    private def self.process_excel_file(filename : String, generator : CroupierGenerator, initial_values : Hash(String, Float64 | String | Bool))
      workbook = ExcelImporter.parse_xlsx(filename)

      workbook.sheets.each do |sheet|
        sheet.cells.each do |cell|
          key = "#{sheet.name}!#{cell.reference}"

          if formula = cell.formula
            generator.add_formula(cell.reference, formula, sheet.name)
          elsif value = cell.value
            # Only add non-nil values
            initial_values[key] = value.as(Float64 | String | Bool)
          end
        end
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

    private def self.print_results(evaluator : Evaluator) : Nil
      cells = evaluator.all_cells

      # Group by sheet
      sheets = Hash(String, Hash(String, Functions::CellValue)).new

      cells.each do |key, value|
        if key.includes?("!")
          parts = key.split("!")
          sheet = parts[0...-1].join("!")
          cell = parts[-1]
        else
          sheet = "(default)"
          cell = key
        end

        sheets[sheet] ||= Hash(String, Functions::CellValue).new
        sheets[sheet][cell] = value
      end

      # Print each sheet
      sheets.each do |sheet_name, sheet_cells|
        puts "--- #{sheet_name} ---"

        # Sort cells by reference (A1, A2, B1, etc.)
        sorted_cells = sheet_cells.to_a.sort do |a, b|
          ref_a = a[0]
          ref_b = b[0]

          # Extract column and row
          col_a = ref_a[/[A-Z]+/]? || ""
          row_a = (ref_a[/\d+$/]? || "0").to_i

          col_b = ref_b[/[A-Z]+/]? || ""
          row_b = (ref_b[/\d+$/]? || "0").to_i

          # Compare by column first, then by row
          if col_a == col_b
            row_a <=> row_b
          else
            col_a <=> col_b
          end
        end

        sorted_cells.each do |ref, value|
          puts "  #{ref}: #{format_value(value)}"
        end

        puts ""
      end
    end

    private def self.format_value(value) : String
      case value
      when Float64
        if value == value.to_i
          value.to_i.to_s
        else
          value.to_s
        end
      when String
        "\"#{value}\""
      when Bool
        value.to_s.upcase
      when Functions::ErrorValue
        value.to_s
      when Nil
        "(empty)"
      when Array
        "[#{value.map { |v| format_value(v) }.join(", ")}]"
      else
        value.to_s
      end
    end

    # Convert Excel file to YAML format and save to examples/
    private def self.convert_excel_to_yaml(filename : String) : String
      # Ensure examples directory exists
      Dir.mkdir("examples") unless Dir.exists?("examples")

      # Get the base filename and change extension to .yaml
      basename = File.basename(filename, ".xlsx")
      yaml_filename = "examples/#{basename}.yaml"

      # If file already exists, append a number
      if File.exists?(yaml_filename)
        counter = 1
        loop do
          new_name = "examples/#{basename}_#{counter}.yaml"
          break unless File.exists?(new_name)
          counter += 1
        end
        yaml_filename = "examples/#{basename}_#{counter}.yaml"
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
              lines << "    #{key}: #{value.inspect}"
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
        Usage: sheety <command> [options]

        Commands:
          sheety <file.(yaml|xlsx)>           Evaluate and print spreadsheet
          sheety compile <file.(yaml|xlsx)>   Compile spreadsheet to standalone binary (interactive mode by default)

        Compile options:
          [output.cr]                 Generated Crystal source file (default: input.(yaml|xlsx) -> input.cr)
          --no-interactive             Generate non-interactive binary (runs once and exits)

        Input formats:
          .xlsx                       Excel 2007+ format (with formula support)
          .yaml, .yml                 YAML format

        YAML Format:
          SheetName:
            A1:
              value: 10
            A2:
              formula: "=SUM(A1:A3)"

        Options:
          -h, --help    Show this help message
      HELP
    end
  end
end
