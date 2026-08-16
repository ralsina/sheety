require "big"
require "yaml"

module Sheety
  # Shared YAML scalar parsing.
  #
  # Both the CLI's spreadsheet reader (Spreadsheet.read_yaml) and the TUI's
  # in-process rebuilder (Rebuilder#process_yaml_data) convert YAML cell
  # values into typed values; this module holds the one true conversion so
  # the two paths can't drift apart again. It deliberately depends on
  # nothing beyond the standard library so the rebuilder can use it inside
  # generated binaries without pulling in the xlsx import stack.
  module YAMLParser
    # Convert a YAML node into a cell value: "true"/"false" strings become
    # Bools, integers and decimals become BigFloats (arbitrary precision,
    # matching how numeric cells are normalized elsewhere), anything else
    # falls back to its string form (Nil becomes "").
    def self.parse_value(value : YAML::Any) : BigFloat | String | Bool
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
      when Int32, Int64, Float64
        # YAML decimals parse as Float64; promote to BigFloat like the TUI
        # does when saving, so numeric cells stay numeric.
        BigFloat.new(raw.to_f, precision: 64)
      when BigFloat
        BigFloat.new(raw, precision: 64)
      when Bool
        raw
      else
        raw.to_s
      end
    end

    # Ensure a spreadsheet UUID is present in raw YAML text: insert it under
    # an existing _ui_state section, append a new section, or — when a UUID
    # is already there — extract it. Returns the effective UUID (an existing
    # one wins over the freshly generated one) and the updated text
    # (identical to the input when a UUID already existed). Shared by
    # Spreadsheet and Rebuilder, which both perform this surgery on YAML
    # files; kept here so the rebuilder doesn't need the xlsx import stack.
    def self.ensure_uuid_in_yaml(yaml_content : String, uuid : String) : Tuple(String, String)
      if yaml_content.includes?("_ui_state:")
        if yaml_content.includes?("spreadsheet_uuid:")
          # Already has it, extract the existing one
          yaml_content.each_line do |line|
            if line.includes?("spreadsheet_uuid:")
              match = line.match(/spreadsheet_uuid:\s*(\S+)/)
              if match
                return {match[1], yaml_content}
              end
            end
          end
          {uuid, yaml_content}
        else
          # Append the UUID to existing _ui_state section
          {uuid, yaml_content.gsub(/(_ui_state:)/, "\\1\n  spreadsheet_uuid: #{uuid}")}
        end
      else
        # Add _ui_state section at the end
        {uuid, yaml_content + "\n_ui_state:\n  spreadsheet_uuid: #{uuid}\n"}
      end
    end
  end
end
