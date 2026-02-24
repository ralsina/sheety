require "../token"
require "../errors"
require "./function_call"

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
          function_token = nil

          while stack.size > 0
            token = stack.pop

            if token.is_a?(Parenthesis) && token.is_opening?
              # Found the matching opening paren
              found_opening = true
              # Check if there's a function token before this paren
              if stack.size > 0 && stack.last.is_a?(Function)
                function_token = stack.pop.as(Function)
              end
              break
            else
              # Collect operator tokens and apply them
              builder.append(token)
            end
          end

          unless found_opening
            raise ParenthesesError.new
          end

          # If we found a function, create the function call
          if function_token
            # Find the opening paren for this function
            # It should be right after the function token
            function_index = tokens.index(function_token)
            opening_paren_index = nil

            if function_index
              # Look for opening paren right after the function
              (function_index + 1).upto(tokens.size - 1) do |i|
                if tokens[i].is_a?(Parenthesis) && tokens[i].as(Parenthesis).is_opening?
                  opening_paren_index = i
                  break
                end
              end
            end

            function_args = Array(AST::Node).new

            if opening_paren_index
              # Get tokens between opening paren and this closing paren
              # Since this closing paren hasn't been added yet, we go to the end
              relevant_tokens = tokens[(opening_paren_index + 1)..-1]

              # Count ArgumentSeparator tokens to determine number of arguments
              separator_count = relevant_tokens.count { |t| t.is_a?(ArgumentSeparator) }
              expected_args = separator_count + 1

              # Count "operand" tokens (things that push nodes to builder)
              operand_count = 0
              relevant_tokens.each do |t|
                if t.is_a?(Operand) || t.is_a?(Function)
                  operand_count += 1
                end
              end

              # Only pop as many nodes as there are operands between the parens
              nodes_to_pop = operand_count
              all_nodes = Array(AST::Node).new
              nodes_to_pop.times do
                if node = builder.pop
                  all_nodes.unshift(node)
                end
              end

              # Distribute nodes into arguments
              if all_nodes.size == expected_args
                # 1:1 mapping
                function_args = all_nodes
              elsif all_nodes.size == 1 && expected_args == 1
                function_args = all_nodes
              else
                # Try to split based on separators
                # For now, simple approach: use all nodes
                function_args = all_nodes
              end
            else
              # No opening paren found - shouldn't happen
              while node = builder.pop
                function_args.unshift(node)
              end
            end

            func_node = function_token.create_function_call(function_args)
            builder.append(func_node)
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
