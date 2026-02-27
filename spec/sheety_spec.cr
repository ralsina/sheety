require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "parsing" do
    it "parses a simple number" do
      ast = Sheety.parse_to_ast("=42")
      ast.should be_a(Sheety::AST::Number)
      ast.expr.should eq("42")
    end

    it "parses TRUE" do
      ast = Sheety.parse_to_ast("=TRUE")
      ast.should be_a(Sheety::AST::Boolean)
      ast.expr.should eq("TRUE")
    end

    it "parses FALSE" do
      ast = Sheety.parse_to_ast("=FALSE")
      ast.should be_a(Sheety::AST::Boolean)
      ast.expr.should eq("FALSE")
    end

    it "parses a float" do
      ast = Sheety.parse_to_ast("=3.14")
      ast.should be_a(Sheety::AST::Number)
      ast.expr.should eq("3.14")
    end

    it "parses scientific notation" do
      ast = Sheety.parse_to_ast("=1E2")
      ast.should be_a(Sheety::AST::Number)
      ast.expr.should eq("100")
    end

    it "detects formulas correctly" do
      # Formulas start with '='
      "=1+2".starts_with?("=").should be_true
      "just text".starts_with?("=").should be_false
      "123".starts_with?("=").should be_false
    end
  end

  describe "AST structure" do
    it "builds AST for simple addition" do
      ast = Sheety.parse_to_ast("=1+2")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("+")
      binop.left.should be_a(Sheety::AST::Number)
      binop.right.should be_a(Sheety::AST::Number)
    end

    it "builds AST with operator precedence" do
      ast = Sheety.parse_to_ast("=1+2*3")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("+")

      # Left should be a number
      binop.left.should be_a(Sheety::AST::Number)

      # Right should be the multiplication
      binop.right.should be_a(Sheety::AST::BinaryOp)
      mult = binop.right.as(Sheety::AST::BinaryOp)
      mult.operator.should eq("*")
    end

    it "builds AST for parentheses" do
      ast = Sheety.parse_to_ast("=(1+2)*3")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("*")

      # Left should be the addition in parens
      binop.left.should be_a(Sheety::AST::BinaryOp)
      add = binop.left.as(Sheety::AST::BinaryOp)
      add.operator.should eq("+")
    end

    it "builds AST for subtraction" do
      ast = Sheety.parse_to_ast("=10-4")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("-")
    end

    it "builds AST for division" do
      ast = Sheety.parse_to_ast("=20/4")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("/")
    end

    it "builds AST for exponentiation" do
      ast = Sheety.parse_to_ast("=2^3")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("^")
    end

    it "builds AST for percent" do
      ast = Sheety.parse_to_ast("=50%")
      ast.should be_a(Sheety::AST::UnaryOp)
      unop = ast.as(Sheety::AST::UnaryOp)
      unop.operator.should eq("%")
    end

    it "builds AST for unary minus" do
      ast = Sheety.parse_to_ast("=-5")
      ast.should be_a(Sheety::AST::UnaryOp)
      unop = ast.as(Sheety::AST::UnaryOp)
      unop.operator.should eq("u-")
      unop.operand.should be_a(Sheety::AST::Number)
    end

    it "builds AST for comparison operators" do
      ast = Sheety.parse_to_ast("=5>3")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq(">")
    end

    it "builds AST for complex expression" do
      ast = Sheety.parse_to_ast("=(1+2)*(3+4)")
      ast.should be_a(Sheety::AST::BinaryOp)
      binop = ast.as(Sheety::AST::BinaryOp)
      binop.operator.should eq("*")
      binop.left.should be_a(Sheety::AST::BinaryOp)
      binop.right.should be_a(Sheety::AST::BinaryOp)
    end

    it "generates correct expression string" do
      ast = Sheety.parse_to_ast("=1+2*3")
      ast.expr.should eq("(1 + (2 * 3))")
    end

    it "generates correct expression with parentheses" do
      ast = Sheety.parse_to_ast("=(1+2)*3")
      ast.expr.should eq("((1 + 2) * 3)")
    end
  end

  describe "errors" do
    it "raises error for mismatched parentheses" do
      expect_raises(Sheety::FormulaError) do
        Sheety.parse_to_ast("=(1+2")
      end
    end
  end
end
