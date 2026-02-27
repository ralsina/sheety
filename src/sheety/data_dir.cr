require "file_utils"
{% unless flag?(:no_embedded_files) %}
  require "./embedded_files"
{% end %}

module Sheety
  # Manages the Sheety data directory
  module DataDir
    # Get the Sheety data directory path
    def self.path : String
      # Use XDG_DATA_HOME or default to ~/.local/share
      if xdg_home = ENV["XDG_DATA_HOME"]?
        File.join(xdg_home, "sheety")
      else
        home = ENV["HOME"]? || raise "HOME environment variable not set"
        File.join(home, ".local", "share", "sheety")
      end
    end

    # Get the sheety binary's modification time
    private def self.sheety_mtime : Time
      sheety_binary = Process.find_executable("sheety")
      raise "sheety binary not found in PATH" unless sheety_binary
      File.info(sheety_binary).modification_time
    end

    # Check if sheety has been updated since data dir was initialized
    private def self.sheety_updated? : Bool
      version_file = File.join(path, ".sheety_version")
      return true unless File.exists?(version_file)
      File.info(version_file).modification_time < sheety_mtime
    end

    # Write version marker after successful update
    private def self.write_version_marker : Nil
      version_file = File.join(path, ".sheety_version")
      File.write(version_file, sheety_mtime.to_unix.to_s)
    end

    # Ensure the data directory exists
    def self.ensure : String
      dir_path = path

      unless Dir.exists?(dir_path)
        # Create parent directories if needed
        FileUtils.mkdir_p(dir_path)
        puts "Created Sheety data directory: #{dir_path}"
      end

      # Also ensure tmp/ directory exists
      tmp_path = File.join(dir_path, "tmp")
      unless Dir.exists?(tmp_path)
        FileUtils.mkdir_p(tmp_path)
      end

      dir_path
    end

    # Ensure shard.yml exists in data directory
    def self.ensure_shard_yml : Nil
      shard_path = File.join(path, "shard.yml")

      # Always recreate shard.yml if sheety was updated
      # This ensures dependencies stay in sync
      if !File.exists?(shard_path) || sheety_updated?
        shard_content = <<-YAML
          name: sheety-generated
          version: 0.1.0

          authors:
            - Sheety User

          targets:
            generated:
              main: generated.cr

          dependencies:
            croupier:
              github: ralsina/croupier
              branch: main
            termisu:
              github: omarluq/termisu
          YAML

        File.write(shard_path, shard_content)
        puts "Created/updated shard.yml in: #{path}"

        # If shard.yml was updated, we need to reinstall dependencies
        # Remove lib/ to force reinstallation
        lib_path = File.join(path, "lib")
        if Dir.exists?(lib_path)
          FileUtils.rm_rf(lib_path)
        end
      end
    end

    # Ensure dependencies are installed (runs shards install if lib/ doesn't exist)
    def self.ensure_dependencies : Nil
      lib_path = File.join(path, "lib")

      unless Dir.exists?(lib_path)
        puts "Installing dependencies in #{path}..."
        status = Process.run("shards", ["install"],
          chdir: path,
          output: Process::Redirect::Inherit,
          error: Process::Redirect::Inherit)

        if status != 0
          STDERR.puts "\nWarning: Failed to install dependencies"
        end
      end
    end

    # Extract embedded source files to data directory (only works in sheety CLI)
    def self.extract_embedded_files : Nil
      src_path = File.join(path, "src")

      # Extract if src doesn't exist OR sheety was updated
      # Only do this if we have EmbeddedFiles available (sheety CLI only)
      if !Dir.exists?(src_path) || sheety_updated?
        {% unless flag?(:no_embedded_files) %}
          puts "Extracting embedded source files..."

          # Remove old src directory if it exists (to clean up deleted files)
          if Dir.exists?(src_path)
            FileUtils.rm_rf(src_path)
          end

          # Iterate through all embedded files and extract them
          Sheety::EmbeddedFiles.files.each do |baked_file|
            # Remove leading slash from baked file path
            baked_path = baked_file.path.lstrip('/')

            # Prepend "src/" to the path since bake_folder strips the root folder name
            file_path = File.join("src", baked_path)

            # Create target path in data directory
            target_path = File.join(path, file_path)

            # Create parent directory if needed
            FileUtils.mkdir_p(File.dirname(target_path))

            # Write the file content
            File.write(target_path, baked_file.gets_to_end)
          end

          # Write version marker after successful extraction
          write_version_marker

          puts "Source files extracted to #{src_path}"
        {% else %}
          # Generated binary: just update version marker to avoid repeated checks
          # The src directory should already exist from when sheety extracted it
          if !Dir.exists?(src_path)
            # If src doesn't exist in generated binary, that's an error - sheety should have created it
            STDERR.puts "Warning: src directory not found. Please run 'sheety <file>' first to initialize."
          end
          write_version_marker
        {% end %}
      end
    end
  end
end
