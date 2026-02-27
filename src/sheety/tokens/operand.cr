require "big"
require "../token"
require "../errors"

module Sheety
  module Tokens
    # Base class for operand tokens (values, references, errors)
    abstract class Operand < Token
      def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
        # Two operands cannot be adjacent
        if tokens.size > 0 && tokens.last.is_a?(Operand)
          raise TokenError.new
        end

        super
        builder.append(self)
        update_n_args(stack)
      end

      protected def update_n_args(stack : Array(Token)) : Nil
        # Default: update operators expecting more arguments
        # This will be used by function calls
      end
    end

    # Numeric literal token
    class Number < Operand
      # Matches integers, floats, scientific notation (not TRUE/FALSE)
      NUMBER_REGEX = /^\s*(?P<name>(?:\d+(?:\.\d+)?|\.\d+)(?:E[+-]?\d+)?)/i

      def self.match?(s : String) : Regex::MatchData?
        NUMBER_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        # Convert string to numeric value using BigFloat
        value_str = name
        @attr["value"] = BigFloat.new(value_str)
        @attr["expr"] = value_str
        @attr["is_boolean"] = false
      end

      def compile : BigFloat | String
        @attr["value"].as(BigFloat)
      end

      def boolean? : Bool
        false
      end
    end

    # Boolean literal token (TRUE, FALSE)
    class Boolean < Operand
      BOOL_REGEX = /^\s*(?P<name>TRUE|FALSE)/i

      def self.match?(s : String) : Regex::MatchData?
        BOOL_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        value_str = name.upcase
        @attr["value"] = (value_str == "TRUE") ? BigFloat.new(1.0, precision: 64) : BigFloat.new(0.0, precision: 64)
        @attr["expr"] = value_str
        @attr["is_boolean"] = true
        @attr["bool_value"] = (value_str == "TRUE")
      end

      def compile : BigFloat | String
        @attr["value"].as(BigFloat)
      end

      def boolean? : Bool
        true
      end

      def bool_value : Bool
        @attr.fetch("bool_value", false).as(Bool)
      end
    end

    # String literal token
    class StringToken < Operand
      # Matches double-quoted strings with escaped quotes
      STRING_REGEX = /^\s*"(?P<name>(?>""|[^""])*)"\s*/

      def self.match?(s : String) : Regex::MatchData?
        STRING_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        # Unescape double quotes
        @attr["expr"] = "\"#{name}\""
        @attr["value"] = name.gsub("\"\"", "\"")
      end

      def compile : BigFloat | String
        @attr["value"].as(String)
      end
    end

    # Excel error values (#DIV/0!, #VALUE!, etc.)
    class ErrorToken < Operand
      ERROR_REGEX = /^\s*(?P<name>#(?:NULL!|DIV\/0!|VALUE!|REF!|NUM!|NAME\?|N\/A))/i

      # Excel error types
      struct XlError
        property value : String

        def initialize(@value : String)
        end

        def to_s(io : IO) : Nil
          io << @value
        end

        def inspect(io : IO) : Nil
          to_s(io)
        end

        # Standard Excel errors
        NULL  = XlError.new("#NULL!")
        DIV   = XlError.new("#DIV/0!")
        VALUE = XlError.new("#VALUE!")
        REF   = XlError.new("#REF!")
        NUM   = XlError.new("#NUM!")
        NAME  = XlError.new("#NAME?")
        NA    = XlError.new("#N/A")

        def self.from_string(s : String) : XlError
          case s.upcase
          when "#NULL!"  then NULL
          when "#DIV/0!" then DIV
          when "#VALUE!" then VALUE
          when "#REF!"   then REF
          when "#NUM!"   then NUM
          when "#NAME?"  then NAME
          when "#N/A"    then NA
          else
            raise FormulaError.new("Unknown error type: #{s}")
          end
        end
      end

      def self.match?(s : String) : Regex::MatchData?
        ERROR_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        @attr["expr"] = name.upcase
        @attr["error"] = XlError.from_string(name).value
      end

      def compile : XlError
        XlError.from_string(name)
      end
    end

    # Cell reference or range token (e.g., A1, Sheet1!B5, A1:B5)
    class Range < Operand
      # Regex for sheet-prefixed full range: Sheet1!A1:B5, 'Sheet 1'!A1:B5
      SHEET_FULL_RANGE_REGEX = /^\s*(?P<sheet>'[^']+'|[A-Z_][A-Z0-9_.]*)!(?P<name>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)/i

      # Regex for sheet-prefixed column range: Sheet1!A:A
      SHEET_COL_RANGE_REGEX = /^\s*(?P<sheet>'[^']+'|[A-Z_][A-Z0-9_.]*)!(?P<name>\$?[A-Z]{1,3}:\$?[A-Z]{1,3})/i

      # Regex for sheet-prefixed row range: Sheet1!1:10
      SHEET_ROW_RANGE_REGEX = /^\s*(?P<sheet>'[^']+'|[A-Z_][A-Z0-9_.]*)!(?P<name>\$?\d+:\$?\d+)/i

      # Regex for sheet-prefixed cell reference: Sheet1!A1, 'Sheet 1'!A1
      SHEET_CELL_REF_REGEX = /^\s*(?P<sheet>'[^']+'|[A-Z_][A-Z0-9_.]*)!(?P<name>\$?[A-Z]{1,3}\$?\d+)/i

      # Regex for full range like A1:B5 or $A$1:$B$5
      FULL_RANGE_REGEX = /^\s*(?P<name>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)/i

      # Regex for column range like A:A
      COL_RANGE_REGEX = /^\s*(?P<name>\$?[A-Z]{1,3}:\$?[A-Z]{1,3})/i

      # Regex for row range like 1:10
      ROW_RANGE_REGEX = /^\s*(?P<name>\$?\d+:\$?\d+)/i

      # Regex for simple cell reference like A1 or $A$1
      CELL_REF_REGEX = /^\s*(?P<name>\$?[A-Z]{1,3}\$?\d+)/i

      def self.match?(s : String) : Regex::MatchData?
        # Try sheet-prefixed references first
        if m = SHEET_FULL_RANGE_REGEX.match(s)
          return m
        end

        if m = SHEET_COL_RANGE_REGEX.match(s)
          return m
        end

        if m = SHEET_ROW_RANGE_REGEX.match(s)
          return m
        end

        if m = SHEET_CELL_REF_REGEX.match(s)
          return m
        end

        # Try full range first (A1:B5)
        if m = FULL_RANGE_REGEX.match(s)
          return m
        end

        # Try column range (A:B)
        if m = COL_RANGE_REGEX.match(s)
          return m
        end

        # Try row range (1:10)
        if m = ROW_RANGE_REGEX.match(s)
          return m
        end

        # Try simple cell reference (A1)
        CELL_REF_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        if m = self.class.match?(s)
          m
        else
          nil
        end
      end

      def process(match : Regex::MatchData) : Regex::MatchData
        @attr["name"] = match["name"]? || match[0]
        super

        # Check for sheet reference
        if sheet = match["sheet"]?
          @attr["sheet"] = sheet
          # Remove quotes from sheet name if present
          if sheet.starts_with?("'") && sheet.ends_with?("'")
            @attr["sheet"] = sheet[1..-2]
          end
        end

        range_str = name
        @attr["expr"] = range_str
        @attr["is_range"] = range_str.includes?(":")
        @attr["is_reference"] = !@attr["is_range"]
        match
      end

      def sheet_name : String?
        @attr["sheet"]?.as(String?)
      end

      def compile : String
        name
      end

      def range? : Bool
        @attr.fetch("is_range", false).as(Bool)
      end

      def reference? : Bool
        @attr.fetch("is_reference", false).as(Bool)
      end
    end

    # Empty operand for missing function arguments
    class EmptyOperand < Operand
      def initialize(@source : String = "", @context : Hash(String, BigFloat | String)? = nil)
        @attr = Hash(String, String | Bool | BigFloat | Int32).new
        @attr["name"] = ""
        @attr["expr"] = ""
        @end_match = 0
      end

      def match(s : String) : Regex::MatchData?
        nil # Never matches - only created programmatically
      end

      def compile : BigFloat | String
        0.0
      end

      def empty? : Bool
        true
      end
    end
  end
end
