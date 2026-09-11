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
require "./pipeline"

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
        return
      end

      # Ensure data directory exists and has required files
      DataDir.ensure
      DataDir.ensure_shard_yml
      DataDir.ensure_dependencies
      DataDir.extract_embedded_files

      # The spreadsheet data comes from the caller-provided intermediate file
      # when there is one (so the TUI's last saved state wins); otherwise from
      # a UUID-keyed intermediate kept in sync with the original file.
      effective_intermediate = @intermediate_file
      unless effective_intermediate
        preliminary_uuid = @spreadsheet_uuid || Spreadsheet.read_with_metadata(filename).uuid
        candidate = Pipeline.intermediate_file(preliminary_uuid)
        if !File.exists?(candidate) || File.info(filename).modification_time > File.info(candidate).modification_time
          FileUtils.cp(filename, candidate)
        end
        effective_intermediate = candidate
      end

      spreadsheet_file = Spreadsheet.read_with_metadata(effective_intermediate)
      spreadsheet_uuid = @spreadsheet_uuid || spreadsheet_file.uuid

      # The content hash (folded with the sheety version) names the generated
      # binary, so an unchanged sheet rebuilds nothing and a changed sheety or
      # sheet gets a fresh binary.
      Pipeline.build(
        spreadsheet_file.data,
        spreadsheet_uuid,
        DataDir.file_hash(@intermediate_file || filename),
        source_file: effective_intermediate,
        original_filename: filename,
      )
    rescue ex : Pipeline::BuildError
      STDERR.puts "\nError: #{ex.message}"
      nil
    end
  end
end
