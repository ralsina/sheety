require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Function calls" do
    it "parses simple function call with no arguments" do
      ast = Sheety.parse_to_ast("=PI()")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("PI")
      func.arguments.size.should eq(0)
    end

    it "parses function call with single argument" do
      ast = Sheety.parse_to_ast("=ABS(A1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("ABS")
      func.arguments.size.should eq(1)
      func.arguments[0].should be_a(Sheety::AST::CellRef)
    end

    it "parses function call with multiple arguments" do
      ast = Sheety.parse_to_ast("=SUM(A1, B1, C1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("SUM")
      func.arguments.size.should eq(3)
    end

    it "parses function call with range argument" do
      ast = Sheety.parse_to_ast("=SUM(A1:B5)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("SUM")
      func.arguments.size.should eq(1)
      func.arguments[0].should be_a(Sheety::AST::RangeRef)
    end

    it "parses function call with numeric argument" do
      ast = Sheety.parse_to_ast("=ROUND(3.14159, 2)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("ROUND")
      func.arguments.size.should eq(2)
    end

    it "parses nested function calls" do
      ast = Sheety.parse_to_ast("=SUM(A1, MAX(B1:B5))")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("SUM")
      func.arguments.size.should eq(2)
      # Second argument should be a nested function call
      func.arguments[1].should be_a(Sheety::AST::FunctionCall)
    end

    it "parses IF function with three arguments" do
      ast = Sheety.parse_to_ast("=IF(A1>0, 1, 0)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("IF")
      func.arguments.size.should eq(3)
    end
  end

  describe "Function names" do
    it "parses uppercase function names" do
      ast = Sheety.parse_to_ast("=SUM(A1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      ast.as(Sheety::AST::FunctionCall).function_name.should eq("SUM")
    end

    it "parses lowercase function names" do
      ast = Sheety.parse_to_ast("=sum(a1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      ast.as(Sheety::AST::FunctionCall).function_name.should eq("sum")
    end

    it "parses mixed case function names" do
      ast = Sheety.parse_to_ast("=Sum(A1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      ast.as(Sheety::AST::FunctionCall).function_name.should eq("Sum")
    end
  end

  describe "Complex formulas with functions" do
    it "parses formula with function and arithmetic" do
      ast = Sheety.parse_to_ast("=SUM(A1:B5) * 2")
      ast.should be_a(Sheety::AST::BinaryOp)
    end

    it "parses formula with nested functions" do
      ast = Sheety.parse_to_ast("=AVERAGE(SUM(A1:A10), SUM(B1:B10))")
      ast.should be_a(Sheety::AST::FunctionCall)
    end
  end
end
