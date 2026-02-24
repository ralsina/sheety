require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Union operator (comma in function arguments)" do
    it "handles multiple range arguments in functions" do
      ast = Sheety.parse_to_ast("=SUM(A1:A5, C1:C5)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.arguments.size.should eq(2)
      func.arguments[0].should be_a(Sheety::AST::RangeRef)
      func.arguments[1].should be_a(Sheety::AST::RangeRef)
    end

    it "handles mixed types in function arguments" do
      ast = Sheety.parse_to_ast("=SUM(A1, 10, B1:B5)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.arguments.size.should eq(3)
    end

    it "handles named ranges as function arguments" do
      ast = Sheety.parse_to_ast("=SUM(MyRange, OtherRange)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.arguments.size.should eq(2)
      func.arguments[0].should be_a(Sheety::AST::NamedRef)
      func.arguments[1].should be_a(Sheety::AST::NamedRef)
    end

    it "handles many arguments" do
      ast = Sheety.parse_to_ast("=SUM(A1, B1, C1, D1, E1)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.arguments.size.should eq(5)
    end
  end
end
