require "../token"
require "../errors"

module Sheety
  module Tokens
    # Named range token (e.g., MyRange, SalesData, Total_Sales)
    class NamedRange < Operand
      # Named ranges start with letter or underscore, followed by letters/numbers/underscore
      # Must NOT be followed by ( to distinguish from function names
      NAME_REGEX = /^\s*(?P<name>[A-Z_][A-Z0-9_.]*)/i

      def self.match?(s : String) : Regex::MatchData?
        if m = NAME_REGEX.match(s)
          name = m["name"]
          rest = s[m.end(0)..-1]

          # Must NOT be followed by ( (that would be a function)
          return nil if rest =~ /^\s*\(/

          # Must NOT be followed by ! (that would be a sheet reference)
          return nil if rest =~ /^\s*!/

          # Must NOT be a true/false
          return nil if name.upcase == "TRUE" || name.upcase == "FALSE"

          # Must NOT be a cell reference pattern: one to three letters (a valid
          # Excel column, ≤ XFD) followed by one or more digits (a row number,
          # up to 1048576). The previous `size <= 3` guard only caught short
          # refs (A1, AB12) and misclassified any 4+ char ref (A100, AA10,
          # CV97, XFD1048576) as a named range — which then dropped it from the
          # dependency graph. Restricting the letter prefix to 1-3 chars keeps
          # legitimate named ranges like Sales2024 (5 letters) working.
          return nil if name =~ /^[A-Z]{1,3}\d+$/i

          # Must NOT be a bare column letter (A, B, ..., XFD) — these are
          # column references, not named ranges.
          return nil if name =~ /^[A-Z]{1,3}$/i

          # Accept multi-letter names and names with underscores
          m
        else
          nil
        end
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        name = match["name"]
        @attr["name"] = name
        @attr["expr"] = name
      end

      def compile : String
        name
      end

      def named_range? : Bool
        true
      end
    end
  end
end
