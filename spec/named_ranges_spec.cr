require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Named ranges" do
    it "parses simple named range" do
      ast = Sheety.parse_to_ast("=MyRange")
      ast.should be_a(Sheety::AST::NamedRef)
      ast.as(Sheety::AST::NamedRef).name.should eq("MyRange")
    end

    it "parses named range with underscores" do
      ast = Sheety.parse_to_ast("=Total_Sales")
      ast.should be_a(Sheety::AST::NamedRef)
      ast.as(Sheety::AST::NamedRef).name.should eq("Total_Sales")
    end

    it "parses multi-word named range" do
      ast = Sheety.parse_to_ast("=SalesData")
      ast.should be_a(Sheety::AST::NamedRef)
    end

    it "distinguishes named range from cell reference" do
      ast = Sheety.parse_to_ast("=MyRange")
      ast.should be_a(Sheety::AST::NamedRef)

      ast2 = Sheety.parse_to_ast("=A1")
      ast2.should be_a(Sheety::AST::CellRef)
    end

    it "parses named range in formula" do
      ast = Sheety.parse_to_ast("=SUM(MyRange, OtherRange)")
      ast.should be_a(Sheety::AST::FunctionCall)
      func = ast.as(Sheety::AST::FunctionCall)
      func.arguments.size.should eq(2)
      func.arguments[0].should be_a(Sheety::AST::NamedRef)
      func.arguments[1].should be_a(Sheety::AST::NamedRef)
    end

    it "parses named range with operators" do
      ast = Sheety.parse_to_ast("=A1 + MyRange")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.left.should be_a(Sheety::AST::CellRef)
      binop.right.should be_a(Sheety::AST::NamedRef)
    end

    it "does not confuse named range with function" do
      # MyRange( would be a function, not a named range
      ast = Sheety.parse_to_ast("=MyRange")
      ast.should be_a(Sheety::AST::NamedRef)

      ast2 = Sheety.parse_to_ast("=MyRange(A1)")
      ast2.should be_a(Sheety::AST::FunctionCall)
    end
  end
end
