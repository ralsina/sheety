require "../token"

module Sheety
  module Tokens
    # Argument separator for function calls (comma)
    # This is different from SeparatorOperator which is for array formulas
    class ArgumentSeparator < Token
      COMMA_REGEX = /^\s*,/

      def self.match?(s : String) : Regex::MatchData?
        COMMA_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        @attr["name"] = ","
        @attr["expr"] = ","
        @end_match = match[0].size
      end

      def compile : String
        ","
      end

      # Argument separators don't participate in operator precedence
      # They just mark argument boundaries
      def lbp : Int32
        0
      end

      def rbp : Int32
        0
      end

      def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
        # Just add to tokens, don't affect the stack or builder
        # The parenthesis closing logic will handle splitting arguments
        tokens << self
      end
    end
  end
end
