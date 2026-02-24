require "./spec_helper"
require "../src/sheety"

describe Sheety do
  describe "Formula parsing (adapted from formulas test_parser.py)" do
    describe "valid formulas" do
      # Basic operators
      it "parses addition" do
        ast = Sheety.parse_to_ast("=1+2")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses subtraction" do
        ast = Sheety.parse_to_ast("=5-3")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses multiplication" do
        ast = Sheety.parse_to_ast("=2*3")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses division" do
        ast = Sheety.parse_to_ast("=10/2")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses exponentiation" do
        ast = Sheety.parse_to_ast("=2^3")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses concatenation" do
        ast = Sheety.parse_to_ast("=\"a\"&\"b\"")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      # Complex expressions
      it "parses parentheses" do
        ast = Sheety.parse_to_ast("=(1+1)+(1+1)")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses nested parentheses" do
        ast = Sheety.parse_to_ast("=((1+1)+1)")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses unary plus" do
        ast = Sheety.parse_to_ast("=+5")
        ast.should be_a(Sheety::AST::UnaryOp)
      end

      it "parses unary minus" do
        ast = Sheety.parse_to_ast("=-5")
        ast.should be_a(Sheety::AST::UnaryOp)
      end

      it "parses percent" do
        ast = Sheety.parse_to_ast("=50%")
        ast.should be_a(Sheety::AST::UnaryOp)
      end

      # Comparison operators
      it "parses less than" do
        ast = Sheety.parse_to_ast("=1<2")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses greater than" do
        ast = Sheety.parse_to_ast("=2>1")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses less than or equal" do
        ast = Sheety.parse_to_ast("=1<=2")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses greater than or equal" do
        ast = Sheety.parse_to_ast("=2>=1")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses equal" do
        ast = Sheety.parse_to_ast("=1=1")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      it "parses not equal" do
        ast = Sheety.parse_to_ast("=1<>2")
        ast.should be_a(Sheety::AST::BinaryOp)
      end

      # Cell references
      it "parses simple cell reference" do
        ast = Sheety.parse_to_ast("=A1")
        ast.should be_a(Sheety::AST::CellRef)
      end

      it "parses column range" do
        ast = Sheety.parse_to_ast("=A:B")
        ast.should be_a(Sheety::AST::RangeRef)
      end

      it "parses row range" do
        ast = Sheety.parse_to_ast("=1:10")
        ast.should be_a(Sheety::AST::RangeRef)
      end

      it "parses cell range" do
        ast = Sheety.parse_to_ast("=A1:B5")
        ast.should be_a(Sheety::AST::RangeRef)
      end

      # Functions
      it "parses function with no arguments" do
        ast = Sheety.parse_to_ast("=PI()")
        ast.should be_a(Sheety::AST::FunctionCall)
        func = ast.as(Sheety::AST::FunctionCall)
        func.arguments.size.should eq(0)
      end

      it "parses function with one argument" do
        ast = Sheety.parse_to_ast("=SUM(A1)")
        ast.should be_a(Sheety::AST::FunctionCall)
      end

      it "parses function with multiple arguments" do
        ast = Sheety.parse_to_ast("=SUM(A1,B1,C1)")
        ast.should be_a(Sheety::AST::FunctionCall)
        func = ast.as(Sheety::AST::FunctionCall)
        func.arguments.size.should eq(3)
      end

      it "parses nested functions" do
        ast = Sheety.parse_to_ast("=SUM(A1,MAX(B1:B5))")
        ast.should be_a(Sheety::AST::FunctionCall)
      end

      # Array constants
      it "parses simple array" do
        ast = Sheety.parse_to_ast("={1,2,3}")
        ast.should be_a(Sheety::AST::ArrayConstant)
      end

      it "parses 2D array" do
        ast = Sheety.parse_to_ast("={1,2;3,4}")
        ast.should be_a(Sheety::AST::ArrayConstant)
      end

      # String literals
      it "parses string literal" do
        ast = Sheety.parse_to_ast("=\"hello\"")
        ast.should be_a(Sheety::AST::StringLiteral)
      end

      it "parses string with spaces" do
        ast = Sheety.parse_to_ast("=\"hello world\"")
        ast.should be_a(Sheety::AST::StringLiteral)
      end

      # Boolean literals
      it "parses TRUE" do
        ast = Sheety.parse_to_ast("=TRUE")
        ast.should be_a(Sheety::AST::Boolean)
        ast.as(Sheety::AST::Boolean).value.should eq(true)
      end

      it "parses FALSE" do
        ast = Sheety.parse_to_ast("=FALSE")
        ast.should be_a(Sheety::AST::Boolean)
        ast.as(Sheety::AST::Boolean).value.should eq(false)
      end

      it "parses mixed case boolean" do
        ast = Sheety.parse_to_ast("=true")
        ast.should be_a(Sheety::AST::Boolean)
      end

      # Error values
      it "parses #NULL!" do
        ast = Sheety.parse_to_ast("=#NULL!")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #DIV/0!" do
        ast = Sheety.parse_to_ast("=#DIV/0!")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #VALUE!" do
        ast = Sheety.parse_to_ast("=#VALUE!")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #REF!" do
        ast = Sheety.parse_to_ast("=#REF!")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #NAME?" do
        ast = Sheety.parse_to_ast("=#NAME?")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #N/A" do
        ast = Sheety.parse_to_ast("=#N/A")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      it "parses #NUM!" do
        ast = Sheety.parse_to_ast("=#NUM!")
        ast.should be_a(Sheety::AST::ErrorValue)
      end

      # Named ranges
      it "parses named range" do
        ast = Sheety.parse_to_ast("=MyRange")
        ast.should be_a(Sheety::AST::NamedRef)
      end

      # Sheet references
      it "parses sheet-prefixed cell" do
        ast = Sheety.parse_to_ast("=Sheet1!A1")
        ast.should be_a(Sheety::AST::CellRef)
        ast.as(Sheety::AST::CellRef).sheet.should eq("Sheet1")
      end

      it "parses quoted sheet name" do
        ast = Sheety.parse_to_ast("='My Sheet'!A1")
        ast.should be_a(Sheety::AST::CellRef)
        ast.as(Sheety::AST::CellRef).sheet.should eq("My Sheet")
      end

      # Operator precedence
      it "handles operator precedence correctly" do
        ast = Sheety.parse_to_ast("=1+2*3")
        ast.should be_a(Sheety::AST::BinaryOp)
        binop = ast.as(Sheety::AST::BinaryOp)
        binop.operator.should eq("+")
        # 2*3 should be evaluated first, so it's the right operand
        binop.right.should be_a(Sheety::AST::BinaryOp)
      end

      it "handles parentheses overriding precedence" do
        ast = Sheety.parse_to_ast("=(1+2)*3")
        ast.should be_a(Sheety::AST::BinaryOp)
        binop = ast.as(Sheety::AST::BinaryOp)
        binop.operator.should eq("*")
        # (1+2) should be evaluated first, so it's the left operand
        binop.left.should be_a(Sheety::AST::BinaryOp)
      end

      # Edge cases
      it "parses single number" do
        ast = Sheety.parse_to_ast("=42")
        ast.should be_a(Sheety::AST::Number)
      end

      it "parses decimal number" do
        ast = Sheety.parse_to_ast("=3.14")
        ast.should be_a(Sheety::AST::Number)
      end

      it "parses scientific notation" do
        ast = Sheety.parse_to_ast("=1E+10")
        ast.should be_a(Sheety::AST::Number)
      end

      it "parses negative number" do
        ast = Sheety.parse_to_ast("=-42")
        ast.should be_a(Sheety::AST::UnaryOp)
      end

      # Complex nested expressions
      it "parses IF function" do
        ast = Sheety.parse_to_ast("=IF(A1>0,1,0)")
        ast.should be_a(Sheety::AST::FunctionCall)
        func = ast.as(Sheety::AST::FunctionCall)
        func.arguments.size.should eq(3)
      end

      it "parses nested IF" do
        ast = Sheety.parse_to_ast("=IF(A1>0,IF(B1>0,1,0),0)")
        ast.should be_a(Sheety::AST::FunctionCall)
      end
    end

    describe "invalid formulas" do
      it "rejects empty formula" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=") }
      end

      it "rejects formula with only operator" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=+") }
      end

      it "rejects formula with mismatched parentheses" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=(1+2") }
      end

      it "rejects formula with extra closing paren" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=1+2)") }
      end

      it "rejects invalid function call" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=SUM(") }
      end

      it "rejects incomplete string" do
        expect_raises(Sheety::FormulaError) { Sheety.parse_to_ast("=\"hello") }
      end
    end
  end
end
