require "../token"
require "../errors"

module Sheety
  module Tokens
    # Array constant token (e.g., {1, 2, 3}, {{1,2},{3,4}})
    class ArrayConstant < Operand
      # Matches array constants enclosed in braces
      # {elements} where elements can be separated by , (column) or ; (row)
      ARRAY_REGEX = /^\s*\{(?P<name>(?>[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}\s*/

      def self.match?(s : String) : Regex::MatchData?
        ARRAY_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        array_str = match["name"]
        @attr["name"] = "{#{array_str}}"
        @attr["expr"] = "{#{array_str}}"
        @end_match = match[0].size
      end

      def compile : String
        name
      end

      def array_content : String
        # Extract content between braces
        if name =~ /^\{(.*)\}$/
          $1
        else
          ""
        end
      end

      # Parse the array content to extract elements
      # This is a simple parser - doesn't handle nested arrays perfectly
      def parse_elements : Array(String)
        content = array_content
        elements = Array(String).new
        current = ""
        depth = 0
        in_string = false

        content.each_char do |char|
          case char
          when '"'
            in_string = !in_string
            current += char
          when '{'
            depth += 1
            current += char
          when '}'
            depth -= 1
            current += char
          when ',', ';'
            if depth == 0 && !in_string
              elements << current.strip unless current.strip.empty?
              current = ""
            else
              current += char
            end
          when ' '
            if !current.strip.empty?
              current += char
            end
          else
            current += char
          end
        end

        # Add last element
        elements << current.strip unless current.strip.empty?

        elements
      end

      def array? : Bool
        true
      end
    end
  end
end
