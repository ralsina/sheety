require "big"
require "openssl"

module Sheety
  # Generates Crystal code from Excel formula AST
  #
  # The generated code is designed to be used in procs that:
  # 1. Read cell values from Croupier's k/v store
  # 2. Calculate the formula result
  # 3. Return the result as a string for storage
  class CodeGenerator
    include AST

    # Context for code generation - tracks sheet name and available cells
    class Context
      property sheet : String?
      property cells : Hash(String, BigFloat | String | Bool)
      # When true, reference leaves render as positional parameters (p0, p1, ...)
      # instead of concrete fetch_cell(...) / fetch_cell_range(...) calls.
      property? parameterized : Bool
      # Counter for the next parameter index, used in parameterized mode.
      property param_index : Int32
      # While visiting the table argument of a lookup function (see
      # TABLE_ARG_POSITIONS), so RangeRef renders as a 2D matrix fetch.
      property table_arg : Bool

      def initialize(@sheet : String? = nil)
        @cells = Hash(String, BigFloat | String | Bool).new
        @parameterized = false
        @param_index = 0
        @table_arg = false
      end
    end

    # Functions that treat one of their arguments as a 2D table (rows and
    # columns of cells) rather than a flat list of values, mapped to the
    # 0-based positions of those arguments. Only direct range references
    # qualify: anything else (a nested expression, an array constant) can't
    # provide the 2D shape the registry functions expect.
    TABLE_ARG_POSITIONS = {"VLOOKUP" => [1], "HLOOKUP" => [1], "INDEX" => [0]}

    # Generate Crystal code for an AST node
    def generate(node : Node, context : Context = Context.new) : String
      visit(node, context)
    end

    # Generate a parameterized expression where each reference leaf (CellRef,
    # RangeRef, NamedRef) is replaced by a positional parameter (p0, p1, ...).
    # Two formulas with the same structure but different cells produce
    # identical parameterized expressions. Used to build shared helper bodies.
    def generate_parameterized(node : Node, sheet : String? = nil) : String
      context = Context.new(sheet)
      context.parameterized = true
      visit(node, context)
    end

    # The ordered list of reference leaves in the AST (CellRef / RangeRef /
    # NamedRef), in the same left-to-right depth-first order the parameterized
    # generator assigns parameter indices. Each entry carries enough info to
    # build the concrete fetch call for a given cell's call site. The optional
    # `sheet` is used to resolve bare references (those without an explicit
    # sheet prefix), mirroring how the concrete generator and dependency
    # extractor resolve them.
    def reference_params(node : Node, sheet : String? = nil) : Array(ReferenceParam)
      params = [] of ReferenceParam
      collect_references(node, params, sheet)
      params
    end

    # A reference leaf extracted for parameterization.
    record ReferenceParam, kind : Symbol, reference : String, sheet : String?

    # Render the concrete fetch expression for a reference parameter, suitable
    # for a task's call site. Mirrors the concrete CellRef/RangeRef rendering.
    def fetch_expression_for(param : ReferenceParam) : String
      case param.kind
      when :cell, :named
        key = param.sheet ? "#{param.sheet}!#{param.reference}" : param.reference
        "fetch_cell(#{key.inspect})"
      when :range, :range2d
        if match = param.reference.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
          sheet = param.sheet
          helper = param.kind == :range2d ? "fetch_cell_range_2d" : "fetch_cell_range"
          "#{helper}(#{sheet.inspect}, #{match[1].inspect}, #{match[2]}, #{match[3].inspect}, #{match[4]})"
        else
          "[]"
        end
      else
        "fetch_cell(#{param.reference.inspect})"
      end
    end

    # The Crystal parameter declaration (name + type) for a reference parameter,
    # used in the shared helper's signature.
    def param_declaration(param : ReferenceParam, index : Int32) : String
      type = case param.kind
             when :range
               "Array(Sheety::Functions::CellValue)"
             when :range2d
               "Array(Array(Sheety::Functions::CellValue))"
             else
               "Sheety::Functions::CellValue"
             end
      "p#{index} : #{type}"
    end

    # A stable structural key for a formula. Two formulas whose ASTs differ
    # only in concrete cell/range/named references share the same key;
    # differences in operators, function names, arity, or literal values
    # produce different keys. Used to group formulas for shared-helper
    # extraction. Returns a hex SHA1 of the canonical structural string.
    def shape_key(node : Node) : String
      canonical = structural_signature(node)
      OpenSSL::Digest.new("SHA1").update(canonical).final.hexstring
    end

    # Visitor methods for each node type

    private def visit(node : Number, context : Context) : String
      "BigFloat.new(#{node.value}, precision: 64)"
    end

    private def visit(node : StringLiteral, context : Context) : String
      node.value.inspect
    end

    private def visit(node : Boolean, context : Context) : String
      node.value.to_s
    end

    private def visit(node : ErrorValue, context : Context) : String
      "Sheety::Functions::ErrorValue.new(#{node.error_value.inspect})"
    end

    private def visit(node : CellRef, context : Context) : String
      return next_param(context) if context.parameterized?

      ref = node.reference.upcase
      sheet = node.sheet || context.sheet

      # Generate code to fetch using helper function
      key = sheet ? "#{sheet}!#{ref}" : ref
      "fetch_cell(#{key.inspect})"
    end

    private def visit(node : RangeRef, context : Context) : String
      return next_param(context) if context.parameterized?

      # Parse range and generate call to fetch_cell_range helper
      range = node.range
      sheet = node.sheet || context.sheet

      # Parse range (e.g., "A1:B5")
      if match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
        start_col = match[1]
        start_row = match[2]
        end_col = match[3]
        end_row = match[4]

        # Table arguments (VLOOKUP et al.) fetch the range as a 2D matrix;
        # every other range fetches a flat, row-major array.
        helper = context.table_arg ? "fetch_cell_range_2d" : "fetch_cell_range"

        # Generate call to helper function
        "#{helper}(#{sheet.inspect}, #{start_col.inspect}, #{start_row}, #{end_col.inspect}, #{end_row})"
      else
        "[]"
      end
    end

    private def visit(node : NamedRef, context : Context) : String
      return next_param(context) if context.parameterized?

      # Named references would need to be resolved beforehand
      # For now, return an error
      "Sheety::Functions::ErrorValue.new(\"#NAME?\")"
    end

    # Allocate the next positional parameter name in parameterized mode.
    private def next_param(context : Context) : String
      name = "p#{context.param_index}"
      context.param_index += 1
      name
    end

    private def visit(node : UnaryOp, context : Context) : String
      operand = visit(node.operand, context)

      case node.operator
      when "+"
        "(+(#{operand}))"
      when "-"
        "(-(#{operand}))"
      when "%"
        "((#{operand}) / 100.0)"
      else
        "(#{node.operator} #{operand})"
      end
    end

    private def visit(node : BinaryOp, context : Context) : String
      left = visit(node.left, context)
      right = visit(node.right, context)

      case node.operator
      when "="
        "Sheety::Functions.eq(#{left}, #{right})"
      when "<>"
        "Sheety::Functions.ne(#{left}, #{right})"
      when "<"
        "Sheety::Functions.lt(#{left}, #{right})"
      when ">"
        "Sheety::Functions.gt(#{left}, #{right})"
      when "<="
        "Sheety::Functions.le(#{left}, #{right})"
      when ">="
        "Sheety::Functions.ge(#{left}, #{right})"
      when "&"
        "Sheety::Functions.to_string(#{left}) + Sheety::Functions.to_string(#{right})"
      when "+"
        "bin_add(#{left}, #{right})"
      when "-"
        "bin_sub(#{left}, #{right})"
      when "*"
        "bin_mul(#{left}, #{right})"
      when "/"
        "bin_div(#{left}, #{right})"
      when "^"
        "bin_pow(#{left}, #{right})"
      else
        "((#{left}) #{node.operator} (#{right}))"
      end
    end

    # Whether the given 0-based argument position of a function call is a 2D
    # table argument backed by a direct range reference.
    private def table_arg?(func_name : String, index : Int32, arg : Node?) : Bool
      return false unless arg.is_a?(RangeRef)
      positions = TABLE_ARG_POSITIONS[func_name]?
      positions ? positions.includes?(index) : false
    end

    # Whether an argument node will be emitted as an array value (a range
    # reference or an array constant), as required by the range parameters
    # of COUNTIF/SUMIF.
    private def array_like?(arg : Node?) : Bool
      arg.is_a?(RangeRef) || arg.is_a?(ArrayConstant)
    end

    private def value_error : String
      "Sheety::Functions::ErrorValue.new(\"#VALUE!\")"
    end

    private def visit(node : FunctionCall, context : Context) : String
      func_name = node.function_name.upcase
      args = node.arguments.map_with_index do |arg, index|
        context.table_arg = table_arg?(func_name, index, arg)
        visit(arg, context)
      end
      context.table_arg = false

      case func_name
      when "SUM"
        "Sheety::Functions.sum(#{args.join(", ")})"
      when "AVERAGE", "AVG"
        "Sheety::Functions.average(#{args.join(", ")})"
      when "MIN"
        "Sheety::Functions.min(#{args.join(", ")})"
      when "MAX"
        "Sheety::Functions.max(#{args.join(", ")})"
      when "COUNT"
        "Sheety::Functions.count(#{args.join(", ")})"
      when "COUNTA"
        "Sheety::Functions.counta(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "MEDIAN"
        "Sheety::Functions.median(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "STDEV", "STDEV.S"
        "Sheety::Functions.stdev(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "STDEV.P"
        "Sheety::Functions.stdev_p(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "VAR.S"
        "Sheety::Functions.var_s(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "VAR.P"
        "Sheety::Functions.var_p(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "ROUND"
        "Sheety::Functions.round(#{args.join(", ")})"
      when "ABS"
        "Sheety::Functions.abs(#{args[0]})"
      when "POWER"
        "Sheety::Functions.power(#{args.join(", ")})"
      when "SQRT"
        "Sheety::Functions.sqrt(#{args[0]})"
      when "MOD"
        "Sheety::Functions.mod(#{args.join(", ")})"
      when "INT"
        "Sheety::Functions.int(#{args[0]})"
      when "CEILING"
        "Sheety::Functions.ceiling(#{args.join(", ")})"
      when "FLOOR"
        "Sheety::Functions.floor(#{args.join(", ")})"
      when "ROUNDUP"
        "Sheety::Functions.roundup(#{args.join(", ")})"
      when "ROUNDDOWN"
        "Sheety::Functions.rounddown(#{args.join(", ")})"
      when "RAND"
        "Sheety::Functions.rand"
      when "RANDBETWEEN"
        "Sheety::Functions.randbetween(#{args.join(", ")})"
      when "IF"
        if args.size >= 3
          "Sheety::Functions.if(#{args[0]}, #{args[1]}, #{args[2]})"
        else
          value_error
        end
      when "IFS"
        # Odd argument counts are rejected (as #VALUE!) by the registry.
        "Sheety::Functions.ifs(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "SWITCH"
        if args.size >= 3
          rest = args[1..]
          if rest.size.odd?
            # Odd remainder: the last argument is the default value.
            "Sheety::Functions.switch_func(#{args[0]}, Sheety::Functions.flatten(#{rest[0...-1].join(", ")}), #{rest.last})"
          else
            "Sheety::Functions.switch_func(#{args[0]}, Sheety::Functions.flatten(#{rest.join(", ")}))"
          end
        else
          value_error
        end
      when "AND"
        "Sheety::Functions.and(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "OR"
        "Sheety::Functions.or(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "NOT"
        "Sheety::Functions.not(#{args[0]})"
      when "CONCAT", "CONCATENATE"
        "Sheety::Functions.concat(Sheety::Functions.flatten(#{args.join(", ")}))"
      when "LEFT"
        "Sheety::Functions.left(#{args.join(", ")})"
      when "RIGHT"
        "Sheety::Functions.right(#{args.join(", ")})"
      when "MID"
        if args.size >= 3
          "Sheety::Functions.mid(#{args.join(", ")})"
        else
          value_error
        end
      when "LEN"
        "Sheety::Functions.len(#{args[0]})"
      when "UPPER"
        "Sheety::Functions.upper(#{args[0]})"
      when "LOWER"
        "Sheety::Functions.lower(#{args[0]})"
      when "TRIM"
        "Sheety::Functions.trim(#{args[0]})"
      when "PROPER"
        "Sheety::Functions.proper(#{args[0]})"
      when "CLEAN"
        "Sheety::Functions.clean(#{args[0]})"
      when "EXACT"
        "Sheety::Functions.exact(#{args.join(", ")})"
      when "REPT"
        "Sheety::Functions.rept(#{args.join(", ")})"
      when "FIND"
        "Sheety::Functions.find(#{args.join(", ")})"
      when "SEARCH"
        "Sheety::Functions.search(#{args.join(", ")})"
      when "SUBSTITUTE"
        "Sheety::Functions.substitute(#{args.join(", ")})"
      when "TEXT"
        "Sheety::Functions.text_func(#{args.join(", ")})"
      when "VALUE"
        "Sheety::Functions.value_func(#{args[0]})"
      when "TODAY"
        "Sheety::Functions.today"
      when "NOW"
        "Sheety::Functions.now"
      when "YEAR"
        "Sheety::Functions.year(#{args[0]})"
      when "MONTH"
        "Sheety::Functions.month(#{args[0]})"
      when "DAY"
        "Sheety::Functions.day(#{args[0]})"
      when "DATEDIF"
        "Sheety::Functions.datedif(#{args.join(", ")})"
      when "EOMONTH"
        "Sheety::Functions.eomonth(#{args.join(", ")})"
      when "COUNTIF"
        if args.size == 2 && array_like?(node.arguments[0]?)
          "Sheety::Functions.countif(#{args.join(", ")})"
        else
          value_error
        end
      when "SUMIF"
        if (2..3).includes?(args.size) && array_like?(node.arguments[0]?) && (args.size == 2 || array_like?(node.arguments[2]?))
          "Sheety::Functions.sumif(#{args.join(", ")})"
        else
          value_error
        end
      when "VLOOKUP"
        if (3..4).includes?(args.size) && table_arg?(func_name, 1, node.arguments[1]?)
          "Sheety::Functions.vlookup(#{args.join(", ")})"
        else
          value_error
        end
      when "HLOOKUP"
        if (3..4).includes?(args.size) && table_arg?(func_name, 1, node.arguments[1]?)
          "Sheety::Functions.hlookup(#{args.join(", ")})"
        else
          value_error
        end
      when "INDEX"
        if (2..3).includes?(args.size) && table_arg?(func_name, 0, node.arguments[0]?)
          "Sheety::Functions.index_func(#{args.join(", ")})"
        else
          value_error
        end
      else
        # Unknown function
        "Sheety::Functions::ErrorValue.new(\"#NAME?\")"
      end
    end

    private def visit(node : ArrayConstant, context : Context) : String
      elements = node.elements.map { |elem| visit(elem, context) }
      "[#{elements.join(", ")}]"
    end

    # Build a canonical structural signature of an AST, where reference leaves
    # (CellRef / RangeRef / NamedRef) all collapse to a single token. This is
    # the basis for the shape key: identical signature => identical structure.
    private def structural_signature(node : Node) : String
      case node
      when CellRef, RangeRef, NamedRef
        # All reference leaves collapse to the same token: the shape depends on
        # *that* there is a reference here, not which cell it points at.
        "ref"
      when Number
        "num(#{node.value})"
      when StringLiteral
        "str(#{node.value})"
      when Boolean
        "bool(#{node.value})"
      when ErrorValue
        "err(#{node.error_value})"
      when UnaryOp
        "un(#{node.operator},#{structural_signature(node.operand)})"
      when BinaryOp
        "bin(#{node.operator},#{structural_signature(node.left)},#{structural_signature(node.right)})"
      when FunctionCall
        args = node.arguments.map { |arg| structural_signature(arg) }.join(",")
        "fn(#{node.function_name.upcase},#{args})"
      when ArrayConstant
        elems = node.elements.map { |elem| structural_signature(elem) }.join(",")
        "arr(#{elems})"
      else
        "unknown"
      end
    end

    # Collect reference leaves in left-to-right depth-first order, matching the
    # order the parameterized generator assigns parameter indices. `table_arg`
    # marks the subtree of a lookup function's table argument so its ranges
    # are declared as 2D parameters.
    private def collect_references(node : Node, params : Array(ReferenceParam), sheet : String?, table_arg : Bool = false) : Nil
      case node
      when CellRef
        params << ReferenceParam.new(:cell, node.reference.upcase, node.sheet || sheet)
      when RangeRef
        params << ReferenceParam.new(table_arg ? :range2d : :range, node.range.upcase, node.sheet || sheet)
      when NamedRef
        params << ReferenceParam.new(:named, node.name, nil)
      when UnaryOp
        collect_references(node.operand, params, sheet, table_arg)
      when BinaryOp
        collect_references(node.left, params, sheet, table_arg)
        collect_references(node.right, params, sheet, table_arg)
      when FunctionCall
        func_name = node.function_name.upcase
        node.arguments.each_with_index do |arg, index|
          collect_references(arg, params, sheet, table_arg?(func_name, index, arg))
        end
      when ArrayConstant
        node.elements.each { |elem| collect_references(elem, params, sheet, table_arg) }
      end
      # Literals (Number, StringLiteral, Boolean, ErrorValue) contribute no params.
    end
  end
end
