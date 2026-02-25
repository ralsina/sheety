require "yaml"
require "./croupier_generator"

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

      yaml_content = File.read(filename)

      # Check for --no-interactive flag to disable interactive mode
      interactive = !extra_args.includes?("--no-interactive")

      # Generate output filename if not specified
      output_cr = extra_args[0]? || filename.gsub(/\.yaml$/, ".cr")

      # Generate the Crystal source file using CroupierGenerator
      generator = CroupierGenerator.new

      # Load YAML and add all formulas and initial values
      data = YAML.parse(yaml_content)
      initial_values = Hash(String, Float64 | String | Bool).new

      data.as_h.each do |sheet_name, sheet_data|
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

      # Generate Croupier task source code with initial values
      source_code = generator.generate_source(initial_values, interactive)

      if source_code.empty?
        STDERR.puts "Error: Failed to generate source code - output is empty"
        exit 1
      end

      # Write the source file
      File.write(output_cr, source_code)

      # Determine binary name
      binary_name = File.basename(output_cr, ".cr")

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

        # Run the binary
        run_result = Process.run("./#{binary_name}", output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

        # If TUI failed (no TTY), show message
        unless run_result.success?
          puts "\nNote: TUI requires a terminal. Run './#{binary_name}' in a terminal to view the spreadsheet."
        end

        exit run_result.exit_code
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

      # Load and process the YAML file
      evaluator = load_yaml_file(filename)

      # Calculate all formulas
      evaluator.calculate_all

      # Print results
      print_results(evaluator)

      exit 0
    end

    private def self.load_yaml_file(filename : String) : Evaluator
      data = YAML.parse(File.read(filename))

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

    private def self.print_help : Nil
      puts <<-HELP
        Usage: sheety <command> [options]

        Commands:
          sheety <file.yaml>           Evaluate and print spreadsheet
          sheety compile <file.yaml>   Compile spreadsheet to standalone binary (interactive mode by default)

        Compile options:
          [output.cr]                 Generated Crystal source file (default: input.yaml -> input.cr)
          --no-interactive             Generate non-interactive binary (runs once and exits)

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
