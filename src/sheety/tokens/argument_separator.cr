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
        # CRITICAL FIX: Process all pending operators on the stack before adding the comma
        # This ensures that expressions like C1>C2 are evaluated before the comma
        while stack.size > 0 && stack.last.is_a?(Operator)
          builder.append(stack.pop)
        end

        # Then add the comma to tokens for later argument counting
        tokens << self
      end
    end
  end
end
