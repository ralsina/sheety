require "spec"
require "../src/sheety"

describe Sheety::CodeGenerator do
  describe "#generate" do
    it "generates code for number literals" do
      ast = Sheety.parse_to_ast("=42")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should eq("42.0")
    end

    it "generates code for string literals" do
      ast = Sheety.parse_to_ast("=\"hello\"")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should eq("\"hello\"")
    end

    it "generates code for boolean literals" do
      ast = Sheety.parse_to_ast("=TRUE")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should eq("true")

      ast = Sheety.parse_to_ast("=FALSE")
      code = gen.generate(ast)
      code.should eq("false")
    end

    it "generates code for binary operations" do
      ast = Sheety.parse_to_ast("=1+2")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("+")
    end

    it "generates code for cell references" do
      ast = Sheety.parse_to_ast("=A1")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("Croupier::TaskManager.get")
      code.should contain("\"A1\"")
    end

    it "generates code for range references" do
      ast = Sheety.parse_to_ast("=SUM(A1:B2)")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("Croupier::TaskManager.get")
    end

    it "generates code for function calls" do
      ast = Sheety.parse_to_ast("=SUM(A1:A5)")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("Sheety::Functions.sum")
    end

    it "generates code for IF function" do
      ast = Sheety.parse_to_ast("=IF(A1>0,1,0)")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("Sheety::Functions.if")
    end

    it "generates code for comparison operators" do
      ast = Sheety.parse_to_ast("=A1=B1")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("Sheety::Functions.eq")
    end

    it "generates code with sheet context" do
      ast = Sheety.parse_to_ast("=A1")
      gen = Sheety::CodeGenerator.new
      context = Sheety::CodeGenerator::Context.new("Sheet1")
      code = gen.generate(ast, context)
      code.should contain("\"Sheet1!A1\"")
    end
  end

  describe "#generate_proc_body" do
    it "generates complete proc body with result conversion" do
      gen = Sheety::CodeGenerator.new
      body = gen.generate_proc_body("=42")
      body.should contain("result")
      body.should contain("case result")
      body.should contain("when Float64")
    end
  end
end

describe Sheety::DependencyExtractor do
  describe "#extract" do
    it "extracts single cell reference" do
      ast = Sheety.parse_to_ast("=A1")
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract(ast)
      deps.should eq(Set{"A1"})
    end

    it "extracts multiple cell references" do
      ast = Sheety.parse_to_ast("=A1+B1")
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract(ast)
      deps.should eq(Set{"A1", "B1"})
    end

    it "extracts range references" do
      ast = Sheety.parse_to_ast("=SUM(A1:B2)")
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract(ast)
      # A1:B2 expands to A1, A2, B1, B2
      deps.should eq(Set{"A1", "A2", "B1", "B2"})
    end

    it "extracts with sheet context" do
      ast = Sheety.parse_to_ast("=A1")
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract(ast, "Sheet1")
      deps.should eq(Set{"Sheet1!A1"})
    end

    it "handles formulas with functions" do
      ast = Sheety.parse_to_ast("=IF(A1>0,SUM(B1:B5),0)")
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract(ast)
      deps.should contain("A1")
      deps.should contain("B1")
      deps.should contain("B5")
    end
  end

  describe "#extract_from_formula" do
    it "extracts dependencies from formula string" do
      extractor = Sheety::DependencyExtractor.new
      deps = extractor.extract_from_formula("=A1+B1")
      deps.should eq(Set{"A1", "B1"})
    end
  end
end

describe Sheety::CroupierGenerator do
  describe "#add_formula" do
    it "adds a formula to the generator" do
      gen = Sheety.croupier_generator
      gen.add_formula("C1", "=SUM(A1:A5)")
      # Should not raise
    end

    it "adds multiple formulas" do
      gen = Sheety.croupier_generator
      gen.add_formulas({"C1" => "=SUM(A1:A5)", "D1" => "=C1*2"})
      # Should not raise
    end

    it "handles formulas with sheet context" do
      gen = Sheety.croupier_generator
      gen.add_formula("C1", "=SUM(A1:A5)", "Sheet1")
      # Should not raise
    end
  end

  describe "#generate_source" do
    it "generates Crystal source code" do
      gen = Sheety.croupier_generator
      gen.add_formula("C1", "=SUM(A1:A5)")
      source = gen.generate_source
      source.should contain("require \"croupier\"")
      source.should contain("Croupier::Task.new")
      source.should contain("SUM")
    end

    it "generates source with dependencies" do
      gen = Sheety.croupier_generator
      gen.add_formula("C1", "=A1+B1")
      source = gen.generate_source
      source.should contain("\"A1\"")
      source.should contain("\"B1\"")
    end
  end
end

describe Sheety::Functions do
  describe "math functions" do
    it "SUM adds numbers" do
      result = Sheety::Functions.sum([1.0, 2.0, 3.0] of Sheety::Functions::CellValue)
      result.should eq(6.0)
    end

    it "AVERAGE calculates mean" do
      result = Sheety::Functions.average([2.0, 4.0] of Sheety::Functions::CellValue)
      result.should eq(3.0)
    end

    it "MIN finds minimum" do
      result = Sheety::Functions.min([1.0, 5.0, 3.0] of Sheety::Functions::CellValue)
      result.should eq(1.0)
    end

    it "MAX finds maximum" do
      result = Sheety::Functions.max([1.0, 5.0, 3.0] of Sheety::Functions::CellValue)
      result.should eq(5.0)
    end

    it "COUNT counts numbers" do
      result = Sheety::Functions.count([1.0, "text", nil, 3.0] of Sheety::Functions::CellValue)
      result.should eq(2.0)
    end

    it "ROUND rounds numbers" do
      result = Sheety::Functions.round(3.14159, 2.0)
      result.should eq(3.14)
    end

    it "ABS returns absolute value" do
      result = Sheety::Functions.abs(-5.0)
      result.should eq(5.0)
    end
  end

  describe "logical functions" do
    it "IF returns true or false branch" do
      result = Sheety::Functions.if(true, "yes", "no")
      result.should eq("yes")

      result = Sheety::Functions.if(false, "yes", "no")
      result.should eq("no")
    end

    it "AND returns true only if all are true" do
      result = Sheety::Functions.and([true, true, true] of Sheety::Functions::CellValue)
      result.should eq(true)

      result = Sheety::Functions.and([true, false, true] of Sheety::Functions::CellValue)
      result.should eq(false)
    end

    it "OR returns true if any is true" do
      result = Sheety::Functions.or([false, true, false] of Sheety::Functions::CellValue)
      result.should eq(true)

      result = Sheety::Functions.or([false, false] of Sheety::Functions::CellValue)
      result.should eq(false)
    end

    it "NOT inverts boolean" do
      result = Sheety::Functions.not(true)
      result.should eq(false)

      result = Sheety::Functions.not(false)
      result.should eq(true)
    end
  end

  describe "text functions" do
    it "CONCAT joins strings" do
      result = Sheety::Functions.concat(["hello", " ", "world"] of Sheety::Functions::CellValue)
      result.should eq("hello world")
    end

    it "LEFT extracts characters from start" do
      result = Sheety::Functions.left("hello", 2.0)
      result.should eq("he")
    end

    it "RIGHT extracts characters from end" do
      result = Sheety::Functions.right("hello", 2.0)
      result.should eq("lo")
    end

    it "MID extracts characters from middle" do
      result = Sheety::Functions.mid("hello", 2.0, 2.0)
      result.should eq("el")
    end

    it "LEN returns string length" do
      result = Sheety::Functions.len("hello")
      result.should eq(5.0)
    end

    it "UPPER converts to uppercase" do
      result = Sheety::Functions.upper("hello")
      result.should eq("HELLO")
    end

    it "LOWER converts to lowercase" do
      result = Sheety::Functions.lower("HELLO")
      result.should eq("hello")
    end
  end

  describe "comparison functions" do
    it "EQ tests equality" do
      Sheety::Functions.eq(1.0, 1.0).should eq(true)
      Sheety::Functions.eq(1.0, 2.0).should eq(false)
    end

    it "NE tests inequality" do
      Sheety::Functions.ne(1.0, 2.0).should eq(true)
      Sheety::Functions.ne(1.0, 1.0).should eq(false)
    end

    it "LT tests less than" do
      Sheety::Functions.lt(1.0, 2.0).should eq(true)
      Sheety::Functions.lt(2.0, 1.0).should eq(false)
    end

    it "GT tests greater than" do
      Sheety::Functions.gt(2.0, 1.0).should eq(true)
      Sheety::Functions.gt(1.0, 2.0).should eq(false)
    end

    it "LE tests less than or equal" do
      Sheety::Functions.le(1.0, 1.0).should eq(true)
      Sheety::Functions.le(1.0, 2.0).should eq(true)
      Sheety::Functions.le(2.0, 1.0).should eq(false)
    end

    it "GE tests greater than or equal" do
      Sheety::Functions.ge(2.0, 2.0).should eq(true)
      Sheety::Functions.ge(2.0, 1.0).should eq(true)
      Sheety::Functions.ge(1.0, 2.0).should eq(false)
    end
  end
end
