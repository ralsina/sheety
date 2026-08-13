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

      def initialize(@sheet : String? = nil)
        @cells = Hash(String, BigFloat | String | Bool).new
        @parameterized = false
        @param_index = 0
      end
    end

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
      when :range
        if match = param.reference.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
          sheet = param.sheet
          "fetch_cell_range(#{sheet.inspect}, #{match[1].inspect}, #{match[2]}, #{match[3].inspect}, #{match[4]})"
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
      type = param.kind == :range ? "Array(Sheety::Functions::CellValue)" : "Sheety::Functions::CellValue"
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

        # Generate call to helper function
        "fetch_cell_range(#{sheet.inspect}, #{start_col.inspect}, #{start_row}, #{end_col.inspect}, #{end_row})"
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

    private def visit(node : FunctionCall, context : Context) : String
      func_name = node.function_name.upcase
      args = node.arguments.map { |arg| visit(arg, context) }

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
      when "IF"
        if args.size >= 3
          "Sheety::Functions.if(#{args[0]}, #{args[1]}, #{args[2]})"
        else
          "Sheety::Functions::ErrorValue.new(\"#VALUE!\")"
        end
      when "AND"
        "Sheety::Functions.and([#{args.join(", ")}])"
      when "OR"
        "Sheety::Functions.or([#{args.join(", ")}])"
      when "NOT"
        "Sheety::Functions.not(#{args[0]})"
      when "CONCAT", "CONCATENATE"
        "Sheety::Functions.concat([#{args.join(", ")}])"
      when "LEFT"
        "Sheety::Functions.left(#{args.join(", ")})"
      when "RIGHT"
        "Sheety::Functions.right(#{args.join(", ")})"
      when "MID"
        if args.size >= 3
          "Sheety::Functions.mid(#{args.join(", ")})"
        else
          "Sheety::Functions::ErrorValue.new(\"#VALUE!\")"
        end
      when "LEN"
        "Sheety::Functions.len(#{args[0]})"
      when "UPPER"
        "Sheety::Functions.upper(#{args[0]})"
      when "LOWER"
        "Sheety::Functions.lower(#{args[0]})"
      when "TRIM"
        "Sheety::Functions.trim(#{args[0]})"
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
    # order the parameterized generator assigns parameter indices.
    private def collect_references(node : Node, params : Array(ReferenceParam), sheet : String?) : Nil
      case node
      when CellRef
        params << ReferenceParam.new(:cell, node.reference.upcase, node.sheet || sheet)
      when RangeRef
        params << ReferenceParam.new(:range, node.range.upcase, node.sheet || sheet)
      when NamedRef
        params << ReferenceParam.new(:named, node.name, nil)
      when UnaryOp
        collect_references(node.operand, params, sheet)
      when BinaryOp
        collect_references(node.left, params, sheet)
        collect_references(node.right, params, sheet)
      when FunctionCall
        node.arguments.each { |arg| collect_references(arg, params, sheet) }
      when ArrayConstant
        node.elements.each { |elem| collect_references(elem, params, sheet) }
      end
      # Literals (Number, StringLiteral, Boolean, ErrorValue) contribute no params.
    end

    # Helper to expand a range like "A1:B2" into cell references
    private def expand_range(start_col : String, start_row : Int32, end_col : String, end_row : Int32, sheet : String?) : Array(String)
      result = [] of String

      # Convert column letters to numbers
      start_col_num = column_to_number(start_col)
      end_col_num = column_to_number(end_col)

      # Iterate through rows and columns
      (start_row..end_row).each do |row|
        (start_col_num..end_col_num).each do |col|
          col_str = number_to_column(col)
          ref = sheet ? "#{sheet}!#{col_str}#{row}" : "#{col_str}#{row}"
          result << ref
        end
      end

      result
    end

    # Convert column letter(s) to number (A=1, Z=26, AA=27, etc.)
    private def column_to_number(col : String) : Int32
      num = 0
      col.each_char { |char| num = num * 26 + (char.ord - 'A'.ord + 1) }
      num
    end

    # Convert column number to letter(s) (1=A, 26=Z, 27=AA, etc.)
    private def number_to_column(num : Int32) : String
      result = ""
      while num > 0
        num -= 1
        result = ('A' + (num % 26)).to_s + result
        num //= 26
      end
      result
    end
  end
end
