require "yaml"

module Sheety
  # YAML parsing utilities
  module YAMLParser
    # Process YAML data and add to generator
    def self.process_yaml_data(data : YAML::Any, generator : CroupierGenerator, initial_values : Hash(String, Float64 | String | Bool)) : Nil
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

    # Parse value from YAML
    def self.parse_value(value : YAML::Any) : Functions::CellValue
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
  end
end
