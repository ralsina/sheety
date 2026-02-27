require "./token"

module Sheety
  # Base class for all tokens in the formula parser
  abstract class Token
    property attr : Hash(String, String | Bool | Float64 | Int32)
    property source : String
    property end_match : Int32

    def initialize(@source : String, @context : Hash(String, Float64 | String)? = nil)
      @attr = Hash(String, String | Bool | Float64 | Int32).new
      @end_match = 0

      match_result = match(@source)
      if match_result
        @end_match = match_result.end(0)
        process(match_result)
      else
        raise TokenError.new(@source)
      end
    end

    # Try to match the token pattern against the input string
    abstract def match(s : String) : Regex::MatchData?

    # Process the match result and populate attributes
    def process(match : Regex::MatchData) : Nil
      match.named_captures.each do |key, value|
        if value
          @attr[key] = value
        end
      end
    end

    # Get the token name from attributes
    def name : String
      @attr.fetch("name", "").as(String)
    end

    # Add this token to the AST - called during parsing
    def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
      tokens << self
    end

    # Update input tokens (used by operators and functions)
    def update_input_tokens(*tokens : Token) : Nil
      # Default implementation does nothing
    end

    # Compile the token to executable code
    def compile : (Array(Float64 | String) -> Float64 | String) | Float64 | String
      raise FormulaError.new("Compile not implemented for #{self.class}")
    end
  end
end
