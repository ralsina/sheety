require "../token"
require "../errors"

module Sheety
  module Tokens
    # Function name token (e.g., SUM, IF, AVERAGE in "SUM(A1:B5)")
    class Function < Token
      # Regex for function names (letters, digits, underscore)
      # Excel function names are typically uppercase letters only
      # NOTE: We don't include the opening paren here - it's matched separately
      FUNCTION_NAME_REGEX = /^\s*(?P<name>[A-Z_][A-Z0-9_.]*)/i

      def self.match?(s : String) : Regex::MatchData?
        # Check if this looks like a function name followed by (
        if m = FUNCTION_NAME_REGEX.match(s)
          name_match = m["name"]
          rest = s[m.end(0)..-1] # Use m.end(0) to get position after the match

          # Only match if the function name is immediately followed by (
          # Skip whitespace and check for opening paren
          if rest =~ /^\s*\(/
            return m
          end
        end
        nil
      end

      def match(s : String) : Regex::MatchData?
        self.class.match?(s)
      end

      def process(match : Regex::MatchData) : Nil
        super
        func_name = match["name"]
        @attr["name"] = func_name
        @attr["function"] = func_name
        @attr["expr"] = func_name
      end

      def compile : String
        @attr["function"].as(String)
      end

      def function_name : String
        @attr.fetch("function", "").as(String)
      end

      # Functions act like left-parentheses in the shunting-yard algorithm
      # They mark the start of a function call
      def lbp : Int32
        8 # Highest precedence
      end

      def rbp : Int32
        8
      end

      def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
        # Push function onto stack to wait for its arguments
        # The closing parenthesis will trigger function call creation
        stack << self
        tokens << self
      end

      # Create the function call AST node with its arguments
      def create_function_call(arguments : Array(AST::Node)) : AST::FunctionCall
        AST::FunctionCall.new(function_name, arguments)
      end
    end
  end
end
