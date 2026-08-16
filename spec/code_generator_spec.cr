require "spec"
require "../src/sheety"

describe Sheety::CodeGenerator do
  describe "#generate" do
    it "generates code for number literals" do
      ast = Sheety.parse_to_ast("=42")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should eq("BigFloat.new(42.0, precision: 64)")
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
      code.should contain("bin_add")
    end

    it "generates code for cell references" do
      ast = Sheety.parse_to_ast("=A1")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("fetch_cell")
      code.should contain("\"A1\"")
    end

    it "generates code for range references" do
      ast = Sheety.parse_to_ast("=SUM(A1:B2)")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("fetch_cell_range")
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

  describe "#generate" do
    it "generates code with fetch_cell helper" do
      ast = Sheety.parse_to_ast("=A1+B1")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("fetch_cell")
    end

    it "generates code with fetch_cell_range helper" do
      ast = Sheety.parse_to_ast("=SUM(A1:A10)")
      gen = Sheety::CodeGenerator.new
      code = gen.generate(ast)
      code.should contain("fetch_cell_range")
    end
  end
end

describe Sheety::CodeGenerator do
  describe "shape deduplication" do
    it "assigns the same shape_key to structurally identical formulas" do
      gen = Sheety::CodeGenerator.new
      left = Sheety.parse_to_ast("=B2*C2")
      right = Sheety.parse_to_ast("=B5*C5")
      gen.shape_key(left).should eq(gen.shape_key(right))
    end

    it "assigns different shape_keys when literal values differ" do
      gen = Sheety::CodeGenerator.new
      with_vat = Sheety.parse_to_ast("=A2*1.21")
      without_vat = Sheety.parse_to_ast("=A2*1.1")
      gen.shape_key(with_vat).should_not eq(gen.shape_key(without_vat))
    end

    it "assigns different shape_keys for different operators" do
      gen = Sheety::CodeGenerator.new
      mul = Sheety.parse_to_ast("=A2*B2")
      add = Sheety.parse_to_ast("=A2+B2")
      gen.shape_key(mul).should_not eq(gen.shape_key(add))
    end

    it "ignores concrete cell references when computing the shape" do
      gen = Sheety::CodeGenerator.new
      sum_a = Sheety.parse_to_ast("=SUM(A1:A3)")
      sum_b = Sheety.parse_to_ast("=SUM(B2:B4)")
      gen.shape_key(sum_a).should eq(gen.shape_key(sum_b))
    end

    it "treats different function names as different shapes" do
      gen = Sheety::CodeGenerator.new
      sum = Sheety.parse_to_ast("=SUM(A1:A3)")
      avg = Sheety.parse_to_ast("=AVERAGE(A1:A3)")
      gen.shape_key(sum).should_not eq(gen.shape_key(avg))
    end

    it "generates a parameterized expression with positional params" do
      gen = Sheety::CodeGenerator.new
      ast = Sheety.parse_to_ast("=B2*C2")
      code = gen.generate_parameterized(ast)
      code.should contain("p0")
      code.should contain("p1")
      # No concrete fetches in the parameterized body.
      code.should_not contain("fetch_cell")
    end

    it "keeps literal values in the parameterized expression" do
      gen = Sheety::CodeGenerator.new
      ast = Sheety.parse_to_ast("=A2*1.21")
      code = gen.generate_parameterized(ast)
      code.should contain("BigFloat.new(1.21")
      code.should contain("p0")
    end

    it "collects reference params in left-to-right order" do
      gen = Sheety::CodeGenerator.new
      ast = Sheety.parse_to_ast("=B2*C2")
      params = gen.reference_params(ast, "Sheet1")
      params.map(&.reference).should eq(["B2", "C2"])
      params.map(&.sheet).should eq(["Sheet1", "Sheet1"])
    end

    it "renders the concrete fetch expression for a cell param" do
      gen = Sheety::CodeGenerator.new
      ast = Sheety.parse_to_ast("=B2")
      param = gen.reference_params(ast, "Sheet1").first
      gen.fetch_expression_for(param).should eq(%(fetch_cell("Sheet1!B2")))
    end

    it "renders the concrete fetch expression for a range param" do
      gen = Sheety::CodeGenerator.new
      ast = Sheety.parse_to_ast("=SUM(B2:B4)")
      param = gen.reference_params(ast, "Sheet1").first
      expr = gen.fetch_expression_for(param)
      expr.should contain("fetch_cell_range")
      expr.should contain(%("Sheet1"))
      expr.should contain(%("B"))
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
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("C1", "=SUM(A1:A5)")
      # Should not raise
    end

    it "handles formulas with sheet context" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("C1", "=SUM(A1:A5)", "Sheet1")
      # Should not raise
    end
  end

  describe "#generate_source" do
    it "generates Crystal source code" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("C1", "=SUM(A1:A5)")
      source = gen.generate_source.entrypoint
      source.should contain("require \"croupier\"")
      # Tasks are registered through a single helper call site, not literal
      # Croupier::Task.new blocks (which OOM'd the compiler on large sheets).
      source.should contain("register_formula_task")
      source.should contain("SUM")
    end

    it "emits exactly one Croupier::Task.new regardless of cell count" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D2", "=B2*C2", "Sheet1")
      gen.add_formula("D3", "=B3*C3", "Sheet1")
      gen.add_formula("D4", "=B4*C4", "Sheet1")
      gen.add_formula("D5", "=B5*C5", "Sheet1")
      source = gen.generate_source.entrypoint

      # The only Croupier::Task.new is the one inside register_formula_task
      # (in croupier_helpers.cr, required by the generated source). The
      # generated body itself must not contain any Task.new literal.
      task_new_count = source.scan("Croupier::Task.new").size
      task_new_count.should eq(0)
      source.should contain("register_formula_task")
    end

    it "generates source with dependencies" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("C1", "=A1+B1")
      source = gen.generate_source.entrypoint
      source.should contain("\"A1\"")
      source.should contain("\"B1\"")
    end

    it "emits one shared calc_shape helper per unique formula shape" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D2", "=B2*C2", "Sheet1")
      gen.add_formula("D3", "=B3*C3", "Sheet1")
      gen.add_formula("D4", "=B4*C4", "Sheet1")
      source = gen.generate_source.entrypoint

      # Exactly one helper definition for the repeated shape.
      shape_defs = source.scan("def calc_shape_").size
      shape_defs.should eq(1)

      # Each task body calls the shared helper rather than inlining arithmetic.
      source.should contain("calc_shape_0(fetch_cell(\"Sheet1!B2\"), fetch_cell(\"Sheet1!C2\"))")
      source.should contain("calc_shape_0(fetch_cell(\"Sheet1!B3\"), fetch_cell(\"Sheet1!C3\"))")
      source.should contain("calc_shape_0(fetch_cell(\"Sheet1!B4\"), fetch_cell(\"Sheet1!C4\"))")
    end

    it "emits separate helpers for distinct shapes" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D2", "=B2*C2", "Sheet1")
      gen.add_formula("E2", "=B2+C2", "Sheet1")
      source = gen.generate_source.entrypoint

      shape_defs = source.scan("def calc_shape_").size
      shape_defs.should eq(2)
    end

    it "preserves task topology (inputs/outputs/id) via the data table" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D2", "=B2*C2", "Sheet1")
      gen.add_formula("D3", "=B3*C3", "Sheet1")
      source = gen.generate_source.entrypoint

      # Each cell still gets its own entry with correct id, inputs, and output.
      source.should contain(%(id: "formula_Sheet1_D2"))
      source.should contain(%(id: "formula_Sheet1_D3"))
      source.should contain(%(["kv://Sheet1!B2", "kv://Sheet1!C2"] of String))
      source.should contain(%(["kv://Sheet1!B3", "kv://Sheet1!C3"] of String))
      source.should contain(%(output: "kv://Sheet1!D2"))
      source.should contain(%(output: "kv://Sheet1!D3"))
    end

    it "does not split small sheets (inline single file)" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D2", "=B2*C2", "Sheet1")
      gen.add_formula("D3", "=B3*C3", "Sheet1")
      generated = gen.generate_source

      generated.split?.should be_false
      generated.aux_files.should be_empty
      # The full table is inlined in the entrypoint.
      generated.entrypoint.should contain("formula_tasks = [")
    end

    it "splits large sheets across chunk files" do
      gen = Sheety::CroupierGenerator.new
      # Exceed SPLIT_THRESHOLD (500) to trigger splitting.
      (2..550).each { |row| gen.add_formula("A#{row}", "=A#{row - 1}+1", "Sheet1") }
      generated = gen.generate_source(chunk_prefix: "testchunk")

      generated.split?.should be_true
      # Two chunk files: 500 + 50 entries.
      generated.aux_files.size.should eq(2)
      generated.aux_files.has_key?("testchunk_tasks_0.cr").should be_true
      generated.aux_files.has_key?("testchunk_tasks_1.cr").should be_true

      # The entrypoint requires the chunks and concatenates their constants.
      generated.entrypoint.should contain(%(require "./testchunk_tasks_0"))
      generated.entrypoint.should contain(%(require "./testchunk_tasks_1"))
      generated.entrypoint.should contain("formula_tasks.concat(TASKS_0)")
      generated.entrypoint.should contain("formula_tasks.concat(TASKS_1)")

      # Each chunk defines its TASKS_N constant.
      generated.aux_files["testchunk_tasks_0.cr"].should contain("TASKS_0 = [")
      generated.aux_files["testchunk_tasks_1.cr"].should contain("TASKS_1 = [")

      # Total task entries across chunks equals the formula count (549).
      total_entries = generated.aux_files.values.sum(&.scan("{id:").size)
      total_entries.should eq(549)
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
      result.should eq(BigFloat.new("3.14"))
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
      result.as(BigFloat).to_f.should be_close(2.138, 0.01)
    end

    it "STDEV.P calculates population standard deviation" do
      result = Sheety::Functions.stdev_p([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(BigFloat).to_f.should be_close(2.0, 0.01)
    end

    it "VAR.S calculates sample variance" do
      result = Sheety::Functions.var_s([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(BigFloat).to_f.should be_close(4.571, 0.01)
    end

    it "VAR.P calculates population variance" do
      result = Sheety::Functions.var_p([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] of Sheety::Functions::CellValue)
      result.as(BigFloat).to_f.should be_close(4.0, 0.01)
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

describe Sheety::CodeGenerator do
  describe "newly wired functions" do
    it "emits flatten-based calls for aggregate functions" do
      {"COUNTA"  => "counta",
       "MEDIAN"  => "median",
       "STDEV"   => "stdev",
       "STDEV.S" => "stdev",
       "STDEV.P" => "stdev_p",
       "VAR.S"   => "var_s",
       "VAR.P"   => "var_p"}.each do |excel, crystal|
        ast = Sheety.parse_to_ast("=#{excel}(A1:A5)")
        code = Sheety::CodeGenerator.new.generate(ast)
        code.should contain("Sheety::Functions.#{crystal}(Sheety::Functions.flatten(fetch_cell_range(")
      end
    end

    it "emits flatten for AND/OR/CONCAT so range arguments compile" do
      {"AND" => "and", "OR" => "or", "CONCAT" => "concat"}.each do |excel, crystal|
        ast = Sheety.parse_to_ast("=#{excel}(A1:A5, B1)")
        code = Sheety::CodeGenerator.new.generate(ast)
        code.should contain("Sheety::Functions.#{crystal}(Sheety::Functions.flatten(fetch_cell_range(")
      end
    end

    it "emits direct calls for passthrough functions" do
      {"CEILING" => "ceiling", "FLOOR" => "floor", "ROUNDUP" => "roundup",
       "ROUNDDOWN" => "rounddown", "RANDBETWEEN" => "randbetween",
       "FIND" => "find", "SEARCH" => "search", "SUBSTITUTE" => "substitute",
       "EXACT" => "exact", "REPT" => "rept", "DATEDIF" => "datedif",
       "EOMONTH" => "eomonth"}.each do |excel, crystal|
        ast = Sheety.parse_to_ast("=#{excel}(A1, 2)")
        code = Sheety::CodeGenerator.new.generate(ast)
        code.should contain("Sheety::Functions.#{crystal}(")
      end
    end

    it "emits renamed registry functions" do
      {"TEXT" => "text_func", "VALUE" => "value_func", "PROPER" => "proper",
       "CLEAN" => "clean", "YEAR" => "year", "MONTH" => "month",
       "DAY" => "day"}.each do |excel, crystal|
        ast = Sheety.parse_to_ast("=#{excel}(A1)")
        code = Sheety::CodeGenerator.new.generate(ast)
        code.should contain("Sheety::Functions.#{crystal}(")
      end
    end

    it "emits bare calls for zero-argument functions, dropping stray args" do
      {"RAND" => "rand", "TODAY" => "today", "NOW" => "now"}.each do |excel, crystal|
        code = Sheety::CodeGenerator.new.generate(Sheety.parse_to_ast("=#{excel}()"))
        code.should eq("Sheety::Functions.#{crystal}")

        code = Sheety::CodeGenerator.new.generate(Sheety.parse_to_ast("=#{excel}(1)"))
        code.should eq("Sheety::Functions.#{crystal}")
      end
    end

    it "emits IFS with flattened condition/value pairs" do
      ast = Sheety.parse_to_ast(%(=IFS(A1>1, "big", A1<0, "small")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("Sheety::Functions.ifs(Sheety::Functions.flatten(")
    end

    it "emits SWITCH with flattened pairs and a separate trailing default" do
      ast = Sheety.parse_to_ast(%(=SWITCH(A1, 1, "one", 2, "two", "other")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("Sheety::Functions.switch_func(fetch_cell(\"A1\"), Sheety::Functions.flatten(")
      code.should contain(%(, "other"))
    end

    it "emits SWITCH without a default for complete pairs" do
      ast = Sheety.parse_to_ast(%(=SWITCH(A1, 1, "one")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should eq(%(Sheety::Functions.switch_func(fetch_cell("A1"), Sheety::Functions.flatten(BigFloat.new(1.0, precision: 64), "one"))))
    end

    it "emits COUNTIF over range arguments" do
      ast = Sheety.parse_to_ast(%(=COUNTIF(A1:A5, ">5")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("Sheety::Functions.countif(fetch_cell_range(")
    end

    it "emits SUMIF with optional sum range" do
      ast = Sheety.parse_to_ast(%(=SUMIF(A1:A5, ">5", B1:B5)))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("Sheety::Functions.sumif(fetch_cell_range(")

      ast = Sheety.parse_to_ast(%(=SUMIF(A1:A5, ">5")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("Sheety::Functions.sumif(fetch_cell_range(")
    end

    it "degrades COUNTIF to #VALUE! for non-array ranges" do
      ast = Sheety.parse_to_ast(%(=COUNTIF(A1, ">5")))
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("#VALUE!")
    end

    it "fetches VLOOKUP table arguments as 2D ranges" do
      ast = Sheety.parse_to_ast("=VLOOKUP(A1, B1:C5, 2)")
      context = Sheety::CodeGenerator::Context.new("Sheet1")
      code = Sheety::CodeGenerator.new.generate(ast, context)
      code.should contain(%(Sheety::Functions.vlookup(fetch_cell("Sheet1!A1"), fetch_cell_range_2d("Sheet1", "B", 1, "C", 5), BigFloat.new(2.0, precision: 64))))
    end

    it "passes VLOOKUP's range_lookup argument through" do
      ast = Sheety.parse_to_ast("=VLOOKUP(A1, B1:C5, 2, FALSE)")
      context = Sheety::CodeGenerator::Context.new("Sheet1")
      code = Sheety::CodeGenerator.new.generate(ast, context)
      code.should contain("Sheety::Functions.vlookup(")
      code.should contain(", false)")
    end

    it "degrades VLOOKUP to #VALUE! when the table is not a range" do
      ast = Sheety.parse_to_ast("=VLOOKUP(A1, D1, 2)")
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("#VALUE!")
    end

    it "degrades VLOOKUP to #VALUE! on wrong arity" do
      ast = Sheety.parse_to_ast("=VLOOKUP(A1, B1:C5)")
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("#VALUE!")
    end

    it "emits INDEX with a 2D table and optional column" do
      ast = Sheety.parse_to_ast("=INDEX(B1:C5, 2)")
      context = Sheety::CodeGenerator::Context.new("Sheet1")
      code = Sheety::CodeGenerator.new.generate(ast, context)
      code.should contain(%(Sheety::Functions.index_func(fetch_cell_range_2d("Sheet1", "B", 1, "C", 5), BigFloat.new(2.0, precision: 64))))
    end

    it "degrades INDEX to #VALUE! when the table is not a range" do
      ast = Sheety.parse_to_ast("=INDEX(A1, 2)")
      code = Sheety::CodeGenerator.new.generate(ast)
      code.should contain("#VALUE!")
    end
  end
end

describe Sheety::CroupierGenerator do
  describe "2D table functions" do
    it "declares 2D parameters for VLOOKUP tables in shared helpers" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D1", "=VLOOKUP(A1, B1:C5, 2)", "Sheet1")
      source = gen.generate_source.entrypoint

      source.should contain("p1 : Array(Array(Sheety::Functions::CellValue))")
      source.should contain(%(calc_shape_0(fetch_cell("Sheet1!A1"), fetch_cell_range_2d("Sheet1", "B", 1, "C", 5))))
    end

    it "initializes and wires inputs for 2D ranges" do
      gen = Sheety::CroupierGenerator.new
      gen.add_formula("D1", "=VLOOKUP(A1, B1:C5, 2)", "Sheet1")
      source = gen.generate_source.entrypoint

      # The setup code must initialize every cell of the table range...
      source.should contain(%(initialize_range("Sheet1", "B", 1, "C", 5)))
      # ...and the task inputs must use the range helper for the same range.
      source.should contain(%(range_inputs("Sheet1", "B", 1, "C", 5)))
    end
  end
end

describe Sheety::Functions do
  describe ".flatten" do
    it "combines scalars and arrays into one flat array" do
      values = Sheety::Functions.flatten("a", ["b", "c"] of Sheety::Functions::CellValue, "d")
      values.should eq(["a", "b", "c", "d"])
    end

    it "returns an empty array for no arguments" do
      Sheety::Functions.flatten.should eq([] of Sheety::Functions::CellValue)
    end
  end

  describe ".index_func" do
    it "defaults the column to 1 for two-argument INDEX" do
      row1 = [1.0.as(Sheety::Functions::CellValue), 2.0.as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      row2 = [3.0.as(Sheety::Functions::CellValue), 4.0.as(Sheety::Functions::CellValue)] of Sheety::Functions::CellValue
      table = [row1, row2] of Array(Sheety::Functions::CellValue)

      Sheety::Functions.index_func(table, 2.0).should eq(3.0)
    end
  end
end

describe Sheety::YAMLParser do
  describe ".parse_value" do
    it "converts integers and decimals to BigFloat" do
      Sheety::YAMLParser.parse_value(YAML.parse("v: 10")["v"]).should eq(BigFloat.new(10.0, precision: 64))
      Sheety::YAMLParser.parse_value(YAML.parse("v: 1.5")["v"]).should eq(BigFloat.new(1.5, precision: 64))
    end

    it "converts booleans and quoted boolean strings" do
      Sheety::YAMLParser.parse_value(YAML.parse("v: true")["v"]).should be_true
      Sheety::YAMLParser.parse_value(YAML.parse("v: \"false\"")["v"]).should be_false
    end

    it "keeps other strings as strings" do
      Sheety::YAMLParser.parse_value(YAML.parse("v: hello")["v"]).should eq("hello")
    end
  end

  describe ".ensure_uuid_in_yaml" do
    it "appends a new _ui_state section when absent" do
      uuid, text = Sheety::YAMLParser.ensure_uuid_in_yaml("Sheet1:\n  A1:\n    value: 1\n", "abc-123")
      uuid.should eq("abc-123")
      text.should contain("_ui_state:\n  spreadsheet_uuid: abc-123")
    end

    it "inserts into an existing _ui_state section" do
      yaml = "Sheet1:\n  A1:\n    value: 1\n_ui_state:\n  active_cell: A1\n"
      uuid, text = Sheety::YAMLParser.ensure_uuid_in_yaml(yaml, "abc-123")
      uuid.should eq("abc-123")
      text.should contain("_ui_state:\n  spreadsheet_uuid: abc-123")
      text.should contain("active_cell: A1")
    end

    it "extracts an existing uuid without rewriting the text" do
      yaml = "_ui_state:\n  spreadsheet_uuid: existing-uuid\n"
      uuid, text = Sheety::YAMLParser.ensure_uuid_in_yaml(yaml, "abc-123")
      uuid.should eq("existing-uuid")
      text.should eq(yaml)
    end
  end
end
