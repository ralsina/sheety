require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Array constants" do
    it "parses simple numeric array" do
      ast = Sheety.parse_to_ast("={1,2,3}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(3)
      arr.elements[0].should be_a(Sheety::AST::Number)
    end

    it "parses array with booleans" do
      ast = Sheety.parse_to_ast("={TRUE,FALSE}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(2)
      arr.elements[0].should be_a(Sheety::AST::Boolean)
    end

    it "parses array with strings" do
      ast = Sheety.parse_to_ast("={\"a\",\"b\",\"c\"}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(3)
      arr.elements[0].should be_a(Sheety::AST::StringLiteral)
    end

    it "parses mixed type array" do
      ast = Sheety.parse_to_ast("={1,\"text\",TRUE}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(3)
      arr.elements[0].should be_a(Sheety::AST::Number)
      arr.elements[1].should be_a(Sheety::AST::StringLiteral)
      arr.elements[2].should be_a(Sheety::AST::Boolean)
    end

    it "parses nested arrays" do
      ast = Sheety.parse_to_ast("={{1,2},{3,4}}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(2)
      arr.elements[0].should be_a(Sheety::AST::ArrayConstant)
      arr.elements[1].should be_a(Sheety::AST::ArrayConstant)
    end

    it "parses array with cell references" do
      ast = Sheety.parse_to_ast("={A1,B1,C1}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(3)
      arr.elements[0].should be_a(Sheety::AST::CellRef)
    end

    it "parses array with semicolon separator" do
      ast = Sheety.parse_to_ast("={1;2;3}")
      ast.should be_a(Sheety::AST::ArrayConstant)
      arr = ast.as(Sheety::AST::ArrayConstant)
      arr.elements.size.should eq(3)
    end
  end
end
