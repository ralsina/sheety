require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Cell references and ranges" do
    it "parses simple cell reference" do
      ast = Sheety.parse_to_ast("=A1")
      ast.should be_a(Sheety::AST::CellRef)
      ast.expr.should eq("A1")
    end

    it "parses cell references with 3+ digit rows as CellRef, not NamedRef" do
      # Regression: refs with 4+ characters (A100, AA10, XFD1048576) were
      # misclassified as NamedRef, which dropped them from the dependency graph
      # and made Croupier raise "Waiting for" on sheets using them.
      ["=A100", "=A1000", "=Z100", "=AA10", "=CV97", "=XFD1048576"].each do |formula|
        ast = Sheety.parse_to_ast(formula)
        ast.should be_a(Sheety::AST::CellRef), formula
      end
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

  describe "CellRefs.parse_range" do
    it "parses a concrete range" do
      bounds = Sheety::CellRefs.parse_range("A1:B5")
      bounds.should eq(Sheety::CellRefs::RangeBounds.new("A", 1, "B", 5))
    end

    it "is case-insensitive and strips $ anchors" do
      bounds = Sheety::CellRefs.parse_range("$a$1:$B$5")
      bounds.should eq(Sheety::CellRefs::RangeBounds.new("A", 1, "B", 5))
    end

    it "normalizes reversed bounds like Excel" do
      Sheety::CellRefs.parse_range("B2:A1").should eq(Sheety::CellRefs::RangeBounds.new("A", 1, "B", 2))
      Sheety::CellRefs.parse_range("C1:A3").should eq(Sheety::CellRefs::RangeBounds.new("A", 1, "C", 3))
      Sheety::CellRefs.parse_range("B5:A1").should eq(Sheety::CellRefs::RangeBounds.new("A", 1, "B", 5))
    end

    it "clamps whole-column ranges to the grid height" do
      bounds = Sheety::CellRefs.parse_range("B:C")
      bounds.should eq(Sheety::CellRefs::RangeBounds.new("B", 1, "C", Sheety::CellRefs::GRID_MAX))
    end

    it "rejects whole-row ranges" do
      Sheety::CellRefs.parse_range("1:10").should be_nil
    end

    it "rejects single references and junk" do
      Sheety::CellRefs.parse_range("A1").should be_nil
      Sheety::CellRefs.parse_range("banana").should be_nil
      Sheety::CellRefs.parse_range("").should be_nil
    end
  end

  describe "CellRefs.parse_ref" do
    it "parses plain and anchored references" do
      Sheety::CellRefs.parse_ref("A1").should eq({col: 1, row: 1})
      Sheety::CellRefs.parse_ref("$B$27").should eq({col: 2, row: 27})
      Sheety::CellRefs.parse_ref("aa10").should eq({col: 27, row: 10})
    end

    it "rejects non-references" do
      Sheety::CellRefs.parse_ref("SUM").should be_nil
      Sheety::CellRefs.parse_ref("A").should be_nil
      Sheety::CellRefs.parse_ref("1").should be_nil
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
