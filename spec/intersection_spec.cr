require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Intersection operator" do
    it "parses simple intersection" do
      ast = Sheety.parse_to_ast("=A:A 1:1")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq(" ")
    end

    it "parses range intersection" do
      ast = Sheety.parse_to_ast("=A1:C3 B2:D4")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq(" ")
      binop.left.should be_a(Sheety::AST::RangeRef)
      binop.right.should be_a(Sheety::AST::RangeRef)
    end

    it "parses intersection with sheet references" do
      ast = Sheety.parse_to_ast("=Sheet1!A:A 1:1")
      ast.should be_a(Sheety::AST::BinaryOp)
    end

    it "parses intersection in function arguments" do
      ast = Sheety.parse_to_ast("=SUM(A:A B:B)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.function_name.should eq("SUM")
    end

    it "handles multiple intersections" do
      ast = Sheety.parse_to_ast("=A:A 1:1 2:2")
      # Should parse as (A:A 1:1) 2:2
      ast.should be_a(Sheety::AST::BinaryOp)
    end
  end
end
