require "big"
require "./errors"
require "./data_dir"
require "./croupier_generator"
require "./spreadsheet"

module Sheety
  # Shared implementation of the "turn spreadsheet data into a compiled
  # binary" flow. The CLI launch path and the TUI's in-process rebuilder
  # previously carried near-identical copies of this (hash -> paths ->
  # generator wiring -> generate -> write -> compile); both now go through
  # here so they can't drift apart again.
  module Pipeline
    # Raised when generation or compilation fails. Callers turn this into a
    # friendly message: the CLI prints it and exits, the rebuilder shows the
    # TUI's "Rebuild failed" notification.
    class BuildError < Exception
    end

    # All on-disk artifacts derived from one spreadsheet identity (content
    # hash + UUID). Hash-named files are the binary and its generated
    # sources; UUID-named files persist across rebuilds.
    record BuildPaths,
      output_cr : String,
      binary_name : String,
      croupier_state : String,
      kv_store : String,
      intermediate_file : String

    def self.paths(file_hash : String, spreadsheet_uuid : String) : BuildPaths
      hash_short = file_hash[0...16]
      BuildPaths.new(
        output_cr: File.join(DataDir.path, "tmp", "#{hash_short}.cr"),
        binary_name: File.join(DataDir.path, "tmp", "#{hash_short}"),
        croupier_state: File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.croupier"),
        kv_store: File.join(DataDir.path, "tmp", "#{spreadsheet_uuid}.kv"),
        intermediate_file: intermediate_file(spreadsheet_uuid),
      )
    end

    # The UUID-keyed intermediate YAML: the copy of the spreadsheet that
    # formula-edit auto-saves go to, so state survives rebuilds.
    def self.intermediate_file(spreadsheet_uuid : String) : String
      File.join(DataDir.path, "#{spreadsheet_uuid}.yaml")
    end

    # Generate the Crystal source for `data` and compile it. Returns the
    # path of the fresh binary.
    #
    # - `source_file` is the YAML the generated TUI restores UI state from
    #   (usually the intermediate file).
    # - `tui_intermediate_file` is where the generated TUI auto-saves formula
    #   edits; nil means the generated program uses `source_file` for that.
    # - `original_filename` is the user-facing file saves go to.
    def self.build(data : Spreadsheet::WorkbookData, spreadsheet_uuid : String, file_hash : String,
                   source_file : String, tui_intermediate_file : String? = nil,
                   original_filename : String? = nil) : String
      build_paths = paths(file_hash, spreadsheet_uuid)

      generator = CroupierGenerator.new
      generator.state_file_path = build_paths.croupier_state
      generator.kv_store_path = build_paths.kv_store
      generator.spreadsheet_uuid = spreadsheet_uuid
      generator.original_filename = original_filename if original_filename

      initial_values = Spreadsheet.populate_generator(data, generator)
      generated = generator.generate_source(
        initial_values, true, source_file, tui_intermediate_file,
        File.basename(build_paths.output_cr, ".cr"),
      )

      if generated.entrypoint.empty?
        raise BuildError.new("Failed to generate source code - output is empty")
      end

      # Write the entrypoint and any chunk files (large sheets split the task
      # table across files to cap compiler memory).
      CroupierGenerator.write_generated(generated, build_paths.output_cr)

      # Build the binary. Run the compiler from the data directory so
      # `require "croupier"` etc. resolve against its lib/ (crystal resolves
      # shard requires from the working directory, which for end users has no
      # lib of its own).
      build_result = Process.run("crystal", ["build", build_paths.output_cr, "-o", build_paths.binary_name],
        chdir: DataDir.path, output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

      unless build_result.success?
        raise BuildError.new("Build failed")
      end

      DataDir.prune_tmp
      build_paths.binary_name
    end
  end
end
