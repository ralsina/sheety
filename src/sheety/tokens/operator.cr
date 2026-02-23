require "../token"
require "../errors"
require "./operand"

module Sheety
  module Tokens
    # Operator token with precedence handling (shunting-yard algorithm)
    abstract class Operator < Token
      # Operator precedence levels (higher = higher precedence)
      PRECEDENCES = {
        ":"  => 8,
        " "  => 8,
        ","  => 8,
        "u-" => 7,
        "u+" => 7,
        "%"  => 6,
        "^"  => 5,
        "*"  => 4,
        "/"  => 4,
        "+"  => 3,
        "-"  => 3,
        "&"  => 2,
        "="  => 1,
        "<"  => 1,
        ">"  => 1,
        "<=" => 1,
        ">=" => 1,
        "<>" => 1,
      }

      # Number of arguments for each operator
      N_ARGS = {
        "u-" => 1,
        "u+" => 1,
        "%"  => 1,
      }

      property operator_name : String

      def initialize(@source : String, @context : Hash(String, Float64 | String)? = nil)
        @operator_name = ""
        super
      end

      def name : String
        @operator_name
      end

      def pred : Int32
        PRECEDENCES[@operator_name]
      end

      def n_args : Int32
        N_ARGS.fetch(@operator_name, 2)
      end

      def ast(tokens : Array(Token), stack : Array(Token), builder : AstBuilder) : Nil
        super
        update_name(tokens)
        current_pred = pred

        # Pop operators with higher or equal precedence
        while stack.size > 0 && stack.last.is_a?(Operator)
          op = stack.last.as(Operator)
          break if current_pred > op.pred
          builder.append(stack.pop)
        end

        stack << self
      end

      protected def update_name(tokens : Array(Token)) : Nil
        # Check for unary operators
        if name == "-" || name == "+"
          # Determine if this is unary or binary
          if should_be_unary?(tokens)
            @operator_name = "u#{name}"
          else
            @operator_name = name
          end
        else
          @operator_name = name
        end
      end

      protected def should_be_unary?(tokens : Array(Token)) : Bool
        return true if tokens.empty?

        last_token = tokens.last
        # Unary if after: opening paren, another operator, or at start
        last_token.is_a?(Parenthesis) && last_token.is_opening? ||
          last_token.is_a?(Operator)
      end

      def update_input_tokens(*input_tokens : Token) : Nil
        # Default: mark as ranges for range operators (:, ,)
        # For basic arithmetic, we don't need special handling
      end

      def to_s(io : IO) : Nil
        io << "#{operator_name} <Operator>"
      end
    end

    # Binary arithmetic operators
    class ArithmeticOperator < Operator
      OPERATOR_REGEX = /^\s*(?P<name>[\+\-\*\/\^])/

      def self.match?(s : String) : Regex::MatchData?
        OPERATOR_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        if m = self.class.match?(s)
          @operator_name = m["name"]
          m
        else
          nil
        end
      end

      def process(match : Regex::MatchData) : Nil
        @attr["name"] = match["name"]
        @operator_name = match["name"]
      end

      def set_expr(*tokens : Token) : Nil
        exprs = tokens.map(&.get_expr)
        @attr["expr"] = "(#{exprs.join(" #{name} ")})"
      end

      # Compile to actual operator function
      def compile : Proc(Array(Float64 | String), Float64 | String)
        case @operator_name
        when "+"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) + args[1].as(Float64)).as(Float64 | String)
          }
        when "-"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) - args[1].as(Float64)).as(Float64 | String)
          }
        when "*"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) * args[1].as(Float64)).as(Float64 | String)
          }
        when "/"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) / args[1].as(Float64)).as(Float64 | String)
          }
        when "^"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) ** args[1].as(Float64)).as(Float64 | String)
          }
        when "u+"
          ->(args : Array(Float64 | String)) {
            (+args[0].as(Float64)).as(Float64 | String)
          }
        when "u-"
          ->(args : Array(Float64 | String)) {
            (-args[0].as(Float64)).as(Float64 | String)
          }
        else
          raise FormulaError.new("Unknown operator: #{@operator_name}")
        end
      end
    end

    # Comparison operators
    class ComparisonOperator < Operator
      COMPARISON_REGEX = /^\s*(?P<name>=|<>|<=|>=|<|>)/

      def self.match?(s : String) : Regex::MatchData?
        COMPARISON_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        if m = self.class.match?(s)
          @operator_name = m["name"]
          m
        else
          nil
        end
      end

      def process(match : Regex::MatchData) : Nil
        @attr["name"] = match["name"]
        @operator_name = match["name"]
      end

      def set_expr(*tokens : Token) : Nil
        exprs = tokens.map(&.get_expr)
        @attr["expr"] = "(#{exprs.join(" #{name} ")})"
      end

      def compile : Proc(Array(Float64 | String), Float64 | String)
        case @operator_name
        when "="
          ->(args : Array(Float64 | String)) {
            (args[0] == args[1] ? 1.0 : 0.0).as(Float64 | String)
          }
        when "<>"
          ->(args : Array(Float64 | String)) {
            (args[0] != args[1] ? 1.0 : 0.0).as(Float64 | String)
          }
        when "<"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) < args[1].as(Float64) ? 1.0 : 0.0).as(Float64 | String)
          }
        when ">"
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) > args[1].as(Float64) ? 1.0 : 0.0).as(Float64 | String)
          }
        when "<="
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) <= args[1].as(Float64) ? 1.0 : 0.0).as(Float64 | String)
          }
        when ">="
          ->(args : Array(Float64 | String)) {
            (args[0].as(Float64) >= args[1].as(Float64) ? 1.0 : 0.0).as(Float64 | String)
          }
        else
          raise FormulaError.new("Unknown comparison operator: #{@operator_name}")
        end
      end
    end

    # Percent operator
    class PercentOperator < Operator
      PERCENT_REGEX = /^\s*(?P<name>%+)/

      def self.match?(s : String) : Regex::MatchData?
        PERCENT_REGEX.match(s)
      end

      def match(s : String) : Regex::MatchData?
        if m = self.class.match?(s)
          @operator_name = "%"
          m
        else
          nil
        end
      end

      def process(match : Regex::MatchData) : Nil
        @attr["name"] = "%"
        @operator_name = "%"
      end

      def set_expr(*tokens : Token) : Nil
        @attr["expr"] = "#{tokens[0].get_expr}%"
      end

      def compile : Proc(Array(Float64 | String), Float64 | String)
        ->(args : Array(Float64 | String)) {
          (args[0].as(Float64) / 100.0).as(Float64 | String)
        }
      end
    end
  end
end
