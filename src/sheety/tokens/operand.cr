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
      # Matches integers, floats, scientific notation, TRUE, FALSE
      NUMBER_REGEX = /^\s*(?P<name>(?:\d+(?:\.\d+)?|\.\d+)(?:E[+-]?\d+)?|TRUE|FALSE)/i

      def self.match?(s : String) : Regex::MatchData?
        NUMBER_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        # Convert string to numeric value
        value_str = name.upcase
        if value_str == "TRUE"
          @attr["value"] = 1.0
        elsif value_str == "FALSE"
          @attr["value"] = 0.0
        else
          @attr["value"] = value_str.to_f
        end
        @attr["expr"] = value_str
      end

      def compile : Float64 | String
        @attr["value"].as(Float64)
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

      def compile : Float64 | String
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
  end
end
