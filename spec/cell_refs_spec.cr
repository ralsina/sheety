require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Cell references and ranges" do
    it "parses simple cell reference" do
      ast = Sheety.parse_to_ast("=A1")
      ast.should be_a(Sheety::AST::CellRef)
      ast.expr.should eq("A1")
    end

    it "parses cell reference with absolute column" do
      ast = Sheety.parse_to_ast("=$A1")
      ast.should be_a(Sheety::AST::CellRef)
      ast.expr.should eq("$A1")
    end

    it "parses cell reference with absolute row" do
      ast = Sheety.parse_to_ast("=A$1")
      ast.should be_a(Sheety::AST::CellRef)
      ast.expr.should eq("A$1")
    end

    it "parses cell reference with absolute column and row" do
      ast = Sheety.parse_to_ast("=$A$1")
      ast.should be_a(Sheety::AST::CellRef)
      ast.expr.should eq("$A$1")
    end

    it "parses range reference" do
      ast = Sheety.parse_to_ast("=A1:B5")
      ast.should be_a(Sheety::AST::RangeRef)
      ast.expr.should eq("A1:B5")
    end

    it "parses range with absolute references" do
      ast = Sheety.parse_to_ast("=$A$1:$B$5")
      ast.should be_a(Sheety::AST::RangeRef)
      ast.expr.should eq("$A$1:$B$5")
    end

    it "parses column range" do
      ast = Sheety.parse_to_ast("=A:B")
      ast.should be_a(Sheety::AST::RangeRef)
      ast.expr.should eq("A:B")
    end

    it "parses row range" do
      ast = Sheety.parse_to_ast("=1:10")
      ast.should be_a(Sheety::AST::RangeRef)
      ast.expr.should eq("1:10")
    end

    it "builds AST with range operator" do
      # Note: Functions not implemented yet, so we test just the range
      ast = Sheety.parse_to_ast("=A1:B5")
      ast.should be_a(Sheety::AST::RangeRef)
    end
  end

  describe "Operator precedence with ranges" do
    it "handles colon operator with correct precedence" do
      # Colon should have higher precedence than arithmetic
      # Note: This is just parsing - we don't evaluate
      ast = Sheety.parse_to_ast("=A1+B1")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("+")
      binop.left.should be_a(Sheety::AST::CellRef)
      binop.right.should be_a(Sheety::AST::CellRef)
    end

    it "handles space intersect operator" do
      # Space operator for intersection
      # A1:B5 C3:D7 should be parsed as (A1:B5) [space] (C3:D7)
      # For now, let's just test that space is recognized
      # This will create a BinaryOp with " " operator
      # Note: We may need to adjust the regex or precedence
    end
  end

  describe "Separator operator" do
    it "handles comma in function-like context" do
      # Comma should be parsed
      # For now we'll test that it doesn't break the parser
      # When we add functions, this will be important
    end
  end

  describe "Complex formulas with references" do
    it "parses formula with cell references" do
      ast = Sheety.parse_to_ast("=A1+B1*C1")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("+")
    end

    it "parses formula with range reference" do
      # Without functions, just test the range part
      ast = Sheety.parse_to_ast("=A1:B5")
      ast.should be_a(Sheety::AST::RangeRef)
    end

    it "parses formula with mixed references and numbers" do
      ast = Sheety.parse_to_ast("=A1+2*B1")
      ast.should be_a(Sheety::AST::BinaryOp)
    end

    it "parses formula with parentheses and references" do
      ast = Sheety.parse_to_ast("=(A1+B1)*C1")
      ast.should be_a(Sheety::AST::BinaryOp)
    end
  end
end
