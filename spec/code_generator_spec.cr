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

  describe "statistical functions" do
    it "COUNTA counts non-empty values" do
      result = Sheety::Functions.counta([1.0, "text", nil, "", 3.0] of Sheety::Functions::CellValue)
      result.should eq(3.0)
    end

    it "COUNTA excludes empty strings" do
      result = Sheety::Functions.counta(["", ""] of Sheety::Functions::CellValue)
      result.should eq(0.0)
    end

    it "MEDIAN finds middle value" do
      result = Sheety::Functions.median([1.0, 3.0, 2.0] of Sheety::Functions::CellValue)
      result.should eq(2.0)
    end

    it "MEDIAN averages middle two values for even count" do
      result = Sheety::Functions.median([1.0, 4.0, 2.0, 3.0] of Sheety::Functions::CellValue)
      result.should eq(2.5)
    end

    it "STDEV calculates sample standard deviation" do
      result = Sheety::Functions.stdev([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(Float64).should be_close(2.138, 0.01)
    end

    it "STDEV.P calculates population standard deviation" do
      result = Sheety::Functions.stdev_p([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(Float64).should be_close(2.0, 0.01)
    end

    it "VAR.S calculates sample variance" do
      result = Sheety::Functions.var_s([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(Float64).should be_close(4.571, 0.01)
    end

    it "VAR.P calculates population variance" do
      result = Sheety::Functions.var_p([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(Float64).should be_close(4.0, 0.01)
    end
  end

  describe "additional math functions" do
    it "CEILING rounds up to nearest multiple" do
      result = Sheety::Functions.ceiling(2.5, 1.0)
      result.should eq(3.0)
      result = Sheety::Functions.ceiling(2.5, 2.0)
      result.should eq(4.0)
    end

    it "FLOOR rounds down to nearest multiple" do
      result = Sheety::Functions.floor(3.7, 2.0)
      result.should eq(2.0)
      result = Sheety::Functions.floor(-3.7, 2.0)
      result.should eq(-4.0)
    end

    it "ROUNDUP rounds away from zero" do
      result = Sheety::Functions.roundup(3.2, 0.0)
      result.should eq(4.0)
      result = Sheety::Functions.roundup(-3.2, 0.0)
      result.should eq(-4.0)
      result = Sheety::Functions.roundup(3.14159, 2.0)
      result.should eq(3.15)
    end

    it "ROUNDDOWN rounds toward zero" do
      result = Sheety::Functions.rounddown(3.7, 0.0)
      result.should eq(3.0)
      result = Sheety::Functions.rounddown(-3.7, 0.0)
      result.should eq(-3.0)
    end

    it "RAND returns number between 0 and 1" do
      result = Sheety::Functions.rand.as(Float64)
      result.should be >= 0.0
      result.should be < 1.0
    end

    it "RANDBETWEEN returns integer in range" do
      result = Sheety::Functions.randbetween(1.0, 10.0).as(Float64)
      result.should be >= 1.0
      result.should be <= 10.0
      result.should eq(result.to_i.to_f)
    end
  end

  describe "additional text functions" do
    it "FIND finds text with case sensitivity" do
      result = Sheety::Functions.find("t", "Text")
      result.should eq(4.0) # lowercase 't' is at position 4
      result = Sheety::Functions.find("T", "Text")
      result.should eq(1.0) # uppercase 'T' is at position 1
    end

    it "SEARCH finds text case-insensitively" do
      result = Sheety::Functions.search("t", "Text")
      result.should eq(1.0)
    end

    it "SUBSTITUTE replaces text" do
      result = Sheety::Functions.substitute("hello world", "world", "there")
      result.should eq("hello there")
    end

    it "SUBSTITUTE replaces nth instance" do
      result = Sheety::Functions.substitute("a a a", "a", "b", 2.0)
      result.should eq("a b a")
    end

    it "TEXT formats number" do
      result = Sheety::Functions.text_func(1234.567, "0.00")
      result.should eq("1234.57")
    end

    it "VALUE converts text to number" do
      result = Sheety::Functions.value_func("123.45")
      result.should eq(123.45)
    end

    it "PROPER capitalizes words" do
      result = Sheety::Functions.proper("hello world")
      result.should eq("Hello World")
    end

    it "CLEAN removes non-printable characters" do
      result = Sheety::Functions.clean("hello\x00world")
      result.should eq("helloworld")
    end

    it "EXACT compares text exactly" do
      Sheety::Functions.exact("hello", "hello").should eq(true)
      Sheety::Functions.exact("hello", "HELLO").should eq(false)
    end

    it "REPT repeats text" do
      result = Sheety::Functions.rept("ab", 3.0)
      result.should eq("ababab")
    end
  end

  describe "date and time functions" do
    it "TODAY returns current date serial" do
      result = Sheety::Functions.today.as(Float64)
      result.should be > 45000.0 # Roughly 2023+
    end

    it "NOW returns current datetime serial" do
      result = Sheety::Functions.now.as(Float64)
      fractional = result % 1.0
      fractional.should be > 0.0 # Should have time component
      fractional.should be < 1.0
    end

    it "YEAR extracts year from date serial" do
      # Date serial for 2023-06-15 is roughly 45098
      result = Sheety::Functions.year(45098.0)
      result.should eq(2023.0)
    end

    it "MONTH extracts month from date serial" do
      result = Sheety::Functions.month(45098.0)
      result.should eq(6.0)
    end

    it "DAY extracts day from date serial" do
      result = Sheety::Functions.day(45098.0)
      result.should eq(21.0) # June 21, 2023
    end

    it "DATEDIF calculates difference in days" do
      # 2023-01-01 to 2023-01-31 = 30 days
      result = Sheety::Functions.datedif(44927.0, 44957.0, "D")
      result.should eq(30.0)
    end

    it "DATEDIF calculates difference in months" do
      result = Sheety::Functions.datedif(44927.0, 45223.0, "M")
      result.should eq(9.0) # Jan to Oct is 9 months
    end

    it "EOMONTH returns last day of month" do
      # Starting from 2023-02-15, EOMONTH should return 2023-02-28
      result = Sheety::Functions.eomonth(44972.0, 0.0)
      result.should eq(44985.0) # Feb 28, 2023
    end
  end

  describe "conditional functions" do
    it "IFS returns first matching value" do
      result = Sheety::Functions.ifs([true, "yes", false, "no"] of Sheety::Functions::CellValue)
      result.should eq("yes")
    end

    it "IFS returns NA if no match" do
      result = Sheety::Functions.ifs([false, "yes", false, "no"] of Sheety::Functions::CellValue)
      result.to_s.should eq("#N/A")
    end

    it "SWITCH returns matching result" do
      result = Sheety::Functions.switch_func(2.0, [1.0, "one", 2.0, "two", 3.0, "three"] of Sheety::Functions::CellValue)
      result.should eq("two")
    end

    it "SWITCH returns default if no match" do
      result = Sheety::Functions.switch_func(4.0, [1.0, "one", 2.0, "two"] of Sheety::Functions::CellValue, "default")
      result.should eq("default")
    end
  end

  describe "conditional aggregation" do
    it "COUNTIF counts cells matching criteria" do
      values = [1.0, 5.0, 3.0, 7.0, 5.0] of Sheety::Functions::CellValue
      result = Sheety::Functions.countif(values, 5.0)
      result.should eq(2.0)
    end

    it "COUNTIF with greater than operator" do
      values = [1.0, 5.0, 3.0, 7.0, 5.0] of Sheety::Functions::CellValue
      result = Sheety::Functions.countif(values, ">4")
      result.should eq(3.0)
    end

    it "SUMIF sums cells matching criteria" do
      range = [1.0, 5.0, 3.0, 7.0, 5.0] of Sheety::Functions::CellValue
      result = Sheety::Functions.sumif(range, 5.0)
      result.should eq(10.0)
    end

    it "SUMIF with operator" do
      range = [1.0, 5.0, 3.0, 7.0, 5.0] of Sheety::Functions::CellValue
      result = Sheety::Functions.sumif(range, ">4")
      result.should eq(17.0) # 5 + 7 + 5
    end

    it "SUMIF with different sum range" do
      range = ["A", "B", "A"] of Sheety::Functions::CellValue
      sum_range = [10.0, 20.0, 30.0] of Sheety::Functions::CellValue
      result = Sheety::Functions.sumif(range, "A", sum_range)
      result.should eq(40.0) # 10 + 30
    end

    it "COUNTIF with wildcard" do
      values = ["apple", "application", "banana"] of Sheety::Functions::CellValue
      result = Sheety::Functions.countif(values, "app*")
      result.should eq(2.0)
    end
  end

  describe "lookup functions" do
    it "VLOOKUP finds exact match" do
      row1 = [1.0.as(Sheety::Functions::CellValue), "One"] of Sheety::Functions::CellValue
      row2 = [2.0.as(Sheety::Functions::CellValue), "Two"] of Sheety::Functions::CellValue
      row3 = [3.0.as(Sheety::Functions::CellValue), "Three"] of Sheety::Functions::CellValue
      table = [row1, row2, row3] of Array(Sheety::Functions::CellValue)

      result = Sheety::Functions.vlookup(2.0, table, 2.0, false)
      result.should eq("Two")
    end

    it "VLOOKUP returns NA for no match" do
      row1 = [1.0.as(Sheety::Functions::CellValue), "One"] of Sheety::Functions::CellValue
      row2 = [2.0.as(Sheety::Functions::CellValue), "Two"] of Sheety::Functions::CellValue
      table = [row1, row2] of Array(Sheety::Functions::CellValue)

      result = Sheety::Functions.vlookup(3.0, table, 2.0, false)
      result.to_s.should eq("#N/A")
    end

    it "HLOOKUP finds value horizontally" do
      row1 = [1.0.as(Sheety::Functions::CellValue), 2.0.as(Sheety::Functions::CellValue), 3.0.as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      row2 = ["One".as(Sheety::Functions::CellValue), "Two".as(Sheety::Functions::CellValue), "Three".as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      table = [row1, row2] of Array(Sheety::Functions::CellValue)

      result = Sheety::Functions.hlookup(2.0, table, 2.0, false)
      result.should eq("Two")
    end

    it "INDEX returns value at position" do
      row1 = [1.0.as(Sheety::Functions::CellValue), 2.0.as(Sheety::Functions::CellValue), 3.0.as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      row2 = [4.0.as(Sheety::Functions::CellValue), 5.0.as(Sheety::Functions::CellValue), 6.0.as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      array = [row1, row2] of Array(Sheety::Functions::CellValue)

      result = Sheety::Functions.index_func(array, 2.0, 2.0)
      result.should eq(5.0)
    end
  end
end
