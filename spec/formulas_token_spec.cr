require "./spec_helper"
require "../src/sheety"

describe Sheety::Tokens do
  describe "Token matching (adapted from formulas test_tokens.py)" do
    describe "Range tokens" do
      it "matches simple cell reference" do
        m = Sheety::Tokens::Range.match?("A1")
        m.should_not be_nil
      end

      it "matches cell reference with absolute column" do
        m = Sheety::Tokens::Range.match?("$A1")
        m.should_not be_nil
      end

      it "matches cell reference with absolute row" do
        m = Sheety::Tokens::Range.match?("A$1")
        m.should_not be_nil
      end

      it "matches cell reference with both absolute" do
        m = Sheety::Tokens::Range.match?("$A$1")
        m.should_not be_nil
      end

      it "matches cell range" do
        m = Sheety::Tokens::Range.match?("A1:B5")
        m.should_not be_nil
      end

      it "matches column range" do
        m = Sheety::Tokens::Range.match?("A:B")
        m.should_not be_nil
      end

      it "matches row range" do
        m = Sheety::Tokens::Range.match?("1:10")
        m.should_not be_nil
      end

      it "matches absolute column range" do
        m = Sheety::Tokens::Range.match?("$A:$B")
        m.should_not be_nil
      end
    end

    describe "Number tokens" do
      it "matches integer" do
        m = Sheety::Tokens::Number.match?("42")
        m.should_not be_nil
      end

      it "matches decimal" do
        m = Sheety::Tokens::Number.match?("3.14")
        m.should_not be_nil
      end

      it "does not match negative decimal (handled by unary minus)" do
        m = Sheety::Tokens::Number.match?("-2.5")
        m.should be_nil
      end

      it "matches scientific notation" do
        m = Sheety::Tokens::Number.match?("1E+10")
        m.should_not be_nil
      end

      it "matches scientific notation negative exponent" do
        m = Sheety::Tokens::Number.match?("1E-5")
        m.should_not be_nil
      end

      it "matches number starting with decimal" do
        m = Sheety::Tokens::Number.match?(".5")
        m.should_not be_nil
      end

      it "does not match TRUE as number" do
        m = Sheety::Tokens::Number.match?("TRUE")
        m.should be_nil
      end

      it "does not match FALSE as number" do
        m = Sheety::Tokens::Number.match?("FALSE")
        m.should be_nil
      end
    end

    describe "Boolean tokens" do
      it "matches TRUE" do
        m = Sheety::Tokens::Boolean.match?("TRUE")
        m.should_not be_nil
      end

      it "matches FALSE" do
        m = Sheety::Tokens::Boolean.match?("FALSE")
        m.should_not be_nil
      end

      it "matches lower case true" do
        m = Sheety::Tokens::Boolean.match?("true")
        m.should_not be_nil
      end

      it "matches mixed case" do
        m = Sheety::Tokens::Boolean.match?("TrUe")
        m.should_not be_nil
      end
    end

    describe "String tokens" do
      it "matches simple string" do
        str = "\"hello\""
        m = Sheety::Tokens::StringToken.match?(str)
        m.should_not be_nil
      end

      it "matches string with escaped quotes" do
        str = "\"hello \"\"world\""
        m = Sheety::Tokens::StringToken.match?(str)
        m.should_not be_nil
      end

      it "matches empty string" do
        str = "\"\""
        m = Sheety::Tokens::StringToken.match?(str)
        m.should_not be_nil
      end

      it "matches string with spaces" do
        str = "\"hello world\""
        m = Sheety::Tokens::StringToken.match?(str)
        m.should_not be_nil
      end
    end

    describe "Error tokens" do
      it "matches #NULL!" do
        m = Sheety::Tokens::ErrorToken.match?("#NULL!")
        m.should_not be_nil
      end

      it "matches #DIV/0!" do
        m = Sheety::Tokens::ErrorToken.match?("#DIV/0!")
        m.should_not be_nil
      end

      it "matches #VALUE!" do
        m = Sheety::Tokens::ErrorToken.match?("#VALUE!")
        m.should_not be_nil
      end

      it "matches #REF!" do
        m = Sheety::Tokens::ErrorToken.match?("#REF!")
        m.should_not be_nil
      end

      it "matches #NAME?" do
        m = Sheety::Tokens::ErrorToken.match?("#NAME?")
        m.should_not be_nil
      end

      it "matches #NUM!" do
        m = Sheety::Tokens::ErrorToken.match?("#NUM!")
        m.should_not be_nil
      end

      it "matches #N/A" do
        m = Sheety::Tokens::ErrorToken.match?("#N/A")
        m.should_not be_nil
      end

      it "matches errors case insensitive" do
        m = Sheety::Tokens::ErrorToken.match?("#null!")
        m.should_not be_nil
      end
    end

    describe "Arithmetic operators" do
      it "matches +" do
        m = Sheety::Tokens::ArithmeticOperator.match?("+")
        m.should_not be_nil
      end

      it "matches -" do
        m = Sheety::Tokens::ArithmeticOperator.match?("-")
        m.should_not be_nil
      end

      it "matches *" do
        m = Sheety::Tokens::ArithmeticOperator.match?("*")
        m.should_not be_nil
      end

      it "matches /" do
        m = Sheety::Tokens::ArithmeticOperator.match?("/")
        m.should_not be_nil
      end

      it "matches ^" do
        m = Sheety::Tokens::ArithmeticOperator.match?("^")
        m.should_not be_nil
      end
    end

    describe "Comparison operators" do
      it "matches =" do
        m = Sheety::Tokens::ComparisonOperator.match?("=")
        m.should_not be_nil
      end

      it "matches <" do
        m = Sheety::Tokens::ComparisonOperator.match?("<")
        m.should_not be_nil
      end

      it "matches >" do
        m = Sheety::Tokens::ComparisonOperator.match?(">")
        m.should_not be_nil
      end

      it "matches <=" do
        m = Sheety::Tokens::ComparisonOperator.match?("<=")
        m.should_not be_nil
      end

      it "matches >=" do
        m = Sheety::Tokens::ComparisonOperator.match?(">=")
        m.should_not be_nil
      end

      it "matches <>" do
        m = Sheety::Tokens::ComparisonOperator.match?("<>")
        m.should_not be_nil
      end
    end

    describe "Concatenation operator" do
      it "matches &" do
        m = Sheety::Tokens::ConcatOperator.match?("&")
        m.should_not be_nil
      end
    end

    describe "Percent operator" do
      it "matches %" do
        m = Sheety::Tokens::PercentOperator.match?("%")
        m.should_not be_nil
      end
    end

    describe "Array constants" do
      it "matches simple array" do
        m = Sheety::Tokens::ArrayConstant.match?("{1,2,3}")
        m.should_not be_nil
      end

      it "matches nested array" do
        m = Sheety::Tokens::ArrayConstant.match?("{{1,2},{3,4}}")
        m.should_not be_nil
      end

      it "matches array with strings" do
        str = "{\"a\",\"b\"}"
        m = Sheety::Tokens::ArrayConstant.match?(str)
        m.should_not be_nil
      end

      it "matches array with booleans" do
        m = Sheety::Tokens::ArrayConstant.match?("{TRUE,FALSE}")
        m.should_not be_nil
      end
    end

    describe "Function names" do
      it "matches function with opening paren" do
        m = Sheety::Tokens::Function.match?("SUM(")
        m.should_not be_nil
      end

      it "matches function with space before paren" do
        m = Sheety::Tokens::Function.match?("SUM (")
        m.should_not be_nil
      end

      it "does not match without paren" do
        m = Sheety::Tokens::Function.match?("SUM")
        m.should be_nil
      end

      it "matches function with numbers in name" do
        m = Sheety::Tokens::Function.match?("SUMIF2(")
        m.should_not be_nil
      end
    end

    describe "Named ranges" do
      it "matches named range" do
        m = Sheety::Tokens::NamedRange.match?("MyRange")
        m.should_not be_nil
      end

      it "matches named range with underscore" do
        m = Sheety::Tokens::NamedRange.match?("Total_Sales")
        m.should_not be_nil
      end

      it "does not match cell reference" do
        m = Sheety::Tokens::NamedRange.match?("A1")
        m.should be_nil
      end

      it "does not match short column names" do
        m = Sheety::Tokens::NamedRange.match?("AB")
        m.should be_nil
      end
    end

    describe "Parenthesis" do
      it "matches opening paren" do
        token = Sheety::Tokens::Parenthesis.new("(")
        m = token.match("(")
        m.should_not be_nil
        token.is_opening?.should eq(true)
      end

      it "matches closing paren" do
        token = Sheety::Tokens::Parenthesis.new(")")
        m = token.match(")")
        m.should_not be_nil
        token.is_opening?.should eq(false)
      end
    end
  end
end
