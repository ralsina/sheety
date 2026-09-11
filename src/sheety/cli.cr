require "big"
require "yaml"
require "docopt-config"
require "./croupier_generator"
require "./spreadsheet"
require "./data_dir"
require "./pipeline"
require "uuid"
require "./importers/excel_importer"

module Sheety
  class CLI
    DOC = <<-DOC
      sheety - compiles spreadsheets into standalone interactive TUI binaries.

      Usage:
        sheety <file> [--save-to=<file>]
        sheety --help | --version

      Options:
        <file>            Spreadsheet to open or convert (.yaml/.yml or .xlsx); created empty if missing
        --save-to=<file>  Convert instead of launching the TUI; format comes from the extension
                          (.yaml/.yml, .xlsx, .cr, .sheety)
        -h --help         Show this help
        --version         Show version

      Spreadsheet YAML format:
        Sheet1:
          A1:
            value: 100
          A2:
            formula: "=SUM(A1:A2)"
      DOC

    def self.run(args : Array(String))
      # Ensure data directory exists on startup
      DataDir.ensure
      DataDir.ensure_shard_yml

      options = parse_args(args)
      filename = options["<file>"].as(String)
      save_to = options["--save-to"]?.as?(String)

      handle_file(filename, save_to)
    rescue ex : Docopt::DocoptExit
      # Usage error: report the problem and the usage line, exit non-zero.
      STDERR.puts ex.message if ex.message
      usage = Docopt::DocoptExit.usage
      STDERR.puts usage unless usage.empty?
      exit 1
    rescue Docopt::ConfigExit
      # --help / --version: the text was already printed to STDOUT.
      exit 0
    rescue ex : Exception
      # Friendly top-level handler: malformed YAML, missing files, failed
      # builds, etc. surface as a one-line error instead of a stack trace.
      STDERR.puts "\nError: #{ex.message}"
      exit 1
    end

    # Public so specs can exercise the arg parsing directly.
    def self.parse_args(args : Array(String)) : Docopt::ConfigOptions
      # env_prefix "" disables environment-variable lookups entirely; sheety
      # has no documented env vars, so none should leak into the options.
      Docopt.docopt_config(DOC, argv: args, exit: false, env_prefix: "",
        version: "sheety #{VERSION}")
    end

    private def self.handle_file(filename : String, save_to : String?) : Nil
      # If file doesn't exist, create an empty spreadsheet
      unless File.exists?(filename)
        puts "Creating new spreadsheet: #{filename}"
        Spreadsheet.create_empty(filename)
      end

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
      source_file = if File.extname(spreadsheet_file.source_file).downcase != ".yaml"
                      temp_yaml = File.join(DataDir.path, "tmp", "#{UUID.random}.yaml")
                      Spreadsheet.write(data, temp_yaml)
                      temp_yaml
                    else
                      spreadsheet_file.source_file
                    end

      # Content hash (folded with the sheety version) for binary naming
      file_hash = DataDir.file_hash(source_file)
      build_paths = Pipeline.paths(file_hash, spreadsheet_uuid)

      # Fast path: if a fresh cached binary already exists (spreadsheet content
      # AND sheety version match, since both are folded into the hash), reuse it
      # directly and skip the heavy data-dir setup (dependency install, embedded
      # file extraction). This avoids redundant shards install/network work on
      # every launch when nothing has changed.
      if File.exists?(build_paths.binary_name) && (!File.exists?(build_paths.output_cr) || File.info(build_paths.binary_name).modification_time >= File.info(build_paths.output_cr).modification_time)
        puts "Using cached binary: #{build_paths.binary_name}"
        puts "\nLaunching TUI..."
        puts "Press Q to exit\n"
        run_result = Process.run(build_paths.binary_name, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)
        unless run_result.success?
          puts "\nNote: TUI requires a terminal. Run '#{build_paths.binary_name}' in a terminal to view the spreadsheet."
        end
        exit run_result.exit_code
      end

      # Cache miss: ensure the build environment is set up before generating
      # and compiling. These are the expensive steps (shards install, file
      # extraction) that the fast path above skips.
      DataDir.ensure_dependencies
      DataDir.extract_embedded_files

      # Intermediate save file for auto-saves (uses UUID to avoid conflicts);
      # refresh it when the source file is newer.
      intermediate_file = build_paths.intermediate_file
      if !File.exists?(intermediate_file) || File.info(source_file).modification_time > File.info(intermediate_file).modification_time
        FileUtils.cp(source_file, intermediate_file)
      end

      puts "Building #{build_paths.binary_name}..."
      binary_name = Pipeline.build(data, spreadsheet_uuid, file_hash,
        source_file: source_file,
        tui_intermediate_file: intermediate_file)

      puts "Built successfully: #{binary_name}"

      puts "\nLaunching TUI..."
      puts "Press Q to exit\n"

      # Run the binary - it handles its own rebuilding via Process.exec
      run_result = Process.run(binary_name, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      # Handle non-zero exit codes
      unless run_result.success?
        puts "\nNote: TUI requires a terminal. Run '#{binary_name}' in a terminal to view the spreadsheet."
      end
      exit run_result.exit_code
    end
  end
end
