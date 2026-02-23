require "../token"
require "../errors"

module Sheety
  module Tokens
    # Parenthesis token for grouping expressions
    class Parenthesis < Token
      PAREN_REGEX = /^(?P<name>[()])/

      property is_opening_paren : Bool = false

      def self.match?(s : String) : Regex::MatchData?
        PAREN_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        if m = self.class.match?(s)
          char = m["name"]
          @is_opening_paren = (char == "(")
          @attr["expr"] = char
          m
        else
          nil
        end
      end

      def process(match : Regex::MatchData) : Nil
        char = match["name"]
        @is_opening_paren = (char == "(")
        @attr["name"] = char
        @attr["expr"] = char
        @attr["is_opening"] = @is_opening_paren
        @attr["is_closing"] = !@is_opening_paren
      end

      def is_opening? : Bool
        @is_opening_paren
      end

      def is_closing? : Bool
        !@is_opening_paren
      end

      def has_start : Bool
        is_opening?
      end

      def has_end : Bool
        is_closing?
      end

      def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
        if is_opening?
          # Push opening paren onto stack
          stack << self
          tokens << self
        else
          # Closing paren: pop until matching opening paren
          found_opening = false
          while stack.size > 0
            token = stack.pop
            if token.is_a?(Parenthesis) && token.is_opening?
              found_opening = true
              break
            end
            builder.append(token)
          end

          unless found_opening
            raise ParenthesesError.new
          end

          tokens << self
        end
      end

      def compile : Float64 | String
        raise FormulaError.new("Parenthesis cannot be compiled directly")
      end
    end
  end
end
