require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "parsing" do
    it "parses a simple number" do
      result = Sheety.evaluate("=42")
      result.should eq(42.0)
    end

    it "parses TRUE" do
      result = Sheety.evaluate("=TRUE")
      result.should eq(1.0)
    end

    it "parses FALSE" do
      result = Sheety.evaluate("=FALSE")
      result.should eq(0.0)
    end

    it "parses a float" do
      result = Sheety.evaluate("=3.14")
      result.should eq(3.14)
    end

    it "parses scientific notation" do
      result = Sheety.evaluate("=1E2")
      result.should eq(100.0)
    end

    it "detects formulas correctly" do
      parser = Sheety::Parser.new
      parser.formula?("=1+2").should be_true
      parser.formula?("just text").should be_false
      parser.formula?("123").should be_false
    end
  end

  describe "errors" do
    it "raises error for invalid formula" do
      expect_raises(Sheety::FormulaError) do
        Sheety.parse("not a formula")
      end
    end
  end
end
