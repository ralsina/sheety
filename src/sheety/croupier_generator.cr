require "big"
require "croupier"
require "./ast"

module Sheety
  # Generates and registers Croupier tasks from Excel formulas
  class CroupierGenerator
    include AST

    # Structure to hold formula information
    struct FormulaInfo
      getter cell : String
      getter formula : String
      getter sheet : String?

      def initialize(@cell : String, @formula : String, @sheet : String? = nil)
      end

      def key : String
        @sheet ? "#{@sheet}!#{@cell.upcase}" : @cell.upcase
      end
    end

    # Result of source generation: the entrypoint source plus any auxiliary
    # chunk files (filename => content) the entrypoint requires. aux_files is
    # empty for small sheets; for large sheets the task table is split across
    # chunk files to keep the Crystal compiler's peak memory low.
    struct GeneratedSource
      getter entrypoint : String
      getter aux_files : Hash(String, String)
      getter? split : Bool

      def initialize(@entrypoint : String, @aux_files : Hash(String, String) = Hash(String, String).new, @split : Bool = false)
      end
    end

    # Above this many formula cells, split the task table across chunk files so
    # the compiler stays within modest memory. Measured: chunk size 500 caps
    # peak RSS at ~1.5GB for a 9900-cell sheet (vs ~13.8GB single-file).
    SPLIT_THRESHOLD = 500
    CHUNK_SIZE      = 500

    # One formula's precomputed artifacts, built once by `build_plans` and
    # shared by the setup, shape-helper and task-registration sections.
    # `ast`/`calc_code` are nil/"" for formulas with a hard problem (they
    # compile to #VALUE! tasks instead).
    private record Plan,
      info : FormulaInfo,
      ast : AST::Node?,
      calc_code : String,
      dependencies : Set(String),
      ranges : Set(DependencyExtractor::RangeDependency),
      problem : String?

    @formulas : Hash(String, FormulaInfo)
    @generator : CodeGenerator
    @extractor : DependencyExtractor
    @state_file_path : String?
    @kv_store_path : String?
    @spreadsheet_uuid : String?
    @original_filename : String?
    # Maps FormulaInfo#key -> index of its shared calc_shape_N helper.
    # Built once per generate_source call; nil outside that scope.
    @shape_assignment : Hash(String, Int32)?
    # Human-readable per-formula problems found by the last generate_source
    # call ("Sheet1!B2: unsupported range ..."). Exposed so callers and
    # specs can assert on what users are warned about.
    getter validation_problems : Array(String)

    def initialize
      @formulas = Hash(String, FormulaInfo).new
      @generator = CodeGenerator.new
      @extractor = DependencyExtractor.new
      @state_file_path = nil
      @kv_store_path = nil
      @spreadsheet_uuid = nil
      @original_filename = nil
      @shape_assignment = nil
      @validation_problems = Array(String).new
    end

    # Set the original filename (for save functionality)
    def original_filename=(filename : String) : self
      @original_filename = filename
      self
    end

    # Set the path for the .croupier state file
    def state_file_path=(path : String) : self
      @state_file_path = path
      self
    end

    # Set the path for the persistent k/v store
    def kv_store_path=(path : String) : self
      @kv_store_path = path
      self
    end

    # Set the spreadsheet UUID (for tracking purposes)
    def spreadsheet_uuid=(uuid : String) : self
      @spreadsheet_uuid = uuid
      self
    end

    # Add a formula to be converted to a task
    def add_formula(cell : String, formula : String, sheet : String? = nil) : self
      # Store formula as-is (we'll strip '=' when parsing)
      key = sheet ? "#{sheet}!#{cell}" : cell

      @formulas[key] = FormulaInfo.new(cell, formula, sheet)
      self
    end

    # Parse formula string and validate it
    private def parse_formula(formula : String) : Node?
      _, builder = Parser.new.ast(formula)
      builder.root
    rescue FormulaError
      nil
    end

    # Sanitize key for use in task ID
    private def sanitize_key(key : String) : String
      key.gsub(/[!\/]/, "_")
    end

    # Parse and lower every formula once, collecting per-formula problems
    # (unparseable formulas, unsupported ranges, oversized ranges, named
    # ranges) instead of failing the whole generation. Hard problems become
    # #VALUE! tasks; named-range formulas keep their (naturally #NAME?)
    # tasks. Everything found lands in validation_problems so callers can
    # warn the user once, in one block.
    private def build_plans : Array(Plan)
      plans = Array(Plan).new
      problems = Array(String).new

      @formulas.each do |_, info|
        formula = info.formula.starts_with?("=") ? info.formula : "=#{info.formula}"

        ast = parse_formula(formula)
        if ast.nil?
          problems << "#{info.key}: #{formula.inspect} could not be parsed; the cell will show #VALUE!"
          plans << Plan.new(info, nil, "", Set(String).new, Set(DependencyExtractor::RangeDependency).new, "unparseable formula")
          next
        end

        # Named ranges parse but have no resolution mechanism; the task will
        # show #NAME?. Warn instead of failing silently.
        named = named_references(ast).uniq
        unless named.empty?
          problems << "#{info.key}: named range#{named.size > 1 ? "s" : ""} #{named.join(", ")} not supported; the cell will show #NAME?"
        end

        begin
          calc_code = @generator.generate(ast, CodeGenerator::Context.new(info.sheet))
          extraction = @extractor.extract_with_ranges(ast, info.sheet)
          plans << Plan.new(info, ast, calc_code, extraction.dependencies, extraction.ranges, nil)
        rescue ex : FormulaError
          problems << "#{info.key}: #{ex.message}; the cell will show #VALUE!"
          plans << Plan.new(info, nil, "", Set(String).new, Set(DependencyExtractor::RangeDependency).new, ex.message || "invalid formula")
        end
      end

      @validation_problems = problems
      plans
    end

    # Collect the names of all named-range references in an AST.
    private def named_references(node : AST::Node) : Array(String)
      names = Array(String).new
      case node
      when AST::NamedRef
        names << node.name
      when AST::UnaryOp
        names.concat(named_references(node.operand))
      when AST::BinaryOp
        names.concat(named_references(node.left))
        names.concat(named_references(node.right))
      when AST::FunctionCall
        node.arguments.each { |argument| names.concat(named_references(argument)) }
      when AST::ArrayConstant
        node.elements.each { |element| names.concat(named_references(element)) }
      end
      names
    end

    # Generate Crystal source code for all tasks.
    # Returns a GeneratedSource: the entrypoint program plus, for large sheets,
    # auxiliary chunk files the entrypoint requires. Callers write the entrypoint
    # to their chosen path and the aux files into the same directory.
    # `chunk_prefix` names the chunk files (e.g. "abc123" -> "abc123_tasks_0.cr")
    # so sheets sharing a tmp dir don't collide; required only when splitting.
    def generate_source(initial_values : Hash(String, BigFloat | String | Bool) = Hash(String, BigFloat | String | Bool).new, interactive : Bool = false, source_file : String? = nil, intermediate_file : String? = nil, chunk_prefix : String = "sheety") : GeneratedSource
      plans = build_plans
      unless @validation_problems.empty?
        STDERR.puts "\nWarning: #{@validation_problems.size} formula problem(s) detected:"
        @validation_problems.each do |problem|
          STDERR.puts "  #{problem}"
        end
      end

      if interactive
        # For interactive mode, require termisu instead of tablo
        source = %(
          require "croupier"
          require "termisu"
          require "../src/sheety/tui"
          require "../src/sheety/functions/registry"
          require "../src/sheety/croupier_helpers"

          # Auto-generated Excel formula tasks for Croupier
          # Generated by Sheety

        )
      else
        source = %(
          require "croupier"
          require "tablo"
          require "../src/sheety/functions/registry"
          require "../src/sheety/croupier_helpers"

          # Auto-generated Excel formula tasks for Croupier
          # Generated by Sheety

        )
      end

      # First add setup code (initial values)
      source += generate_setup_code(initial_values, plans)
      source += "\n\n"

      # Ensure directories exist for state files
      if @state_file_path || @kv_store_path
        source += "# Ensure parent directories exist for state files\n"
        if @state_file_path
          source += "state_dir = File.dirname(#{@state_file_path.inspect})\n"
          source += "Dir.mkdir_p(state_dir) unless Dir.exists?(state_dir)\n"
        end
        if @kv_store_path
          source += "Dir.mkdir_p(#{@kv_store_path.inspect}) unless Dir.exists?(#{@kv_store_path.inspect})\n"
        end
        source += "\n"
      end

      # Configure state file path if set
      if @state_file_path
        source += "# Configure Croupier state file path\n"
        source += "Croupier::TaskManager.state_file = #{@state_file_path.inspect}\n\n"
      end

      # Configure persistent k/v store if set
      if @kv_store_path
        source += "# Configure persistent k/v store for caching results across runs\n"
        source += "Croupier::TaskManager.use_persistent_store(#{@kv_store_path.inspect})\n\n"
      end

      source += "\n"

      # Build shape assignment and emit one shared helper per unique formula
      # shape. Each parseable formula's task body becomes a call to its helper,
      # so a formula repeated across many cells only generates its calculation
      # logic once.
      source += generate_shape_helpers(plans)
      source += "\n\n"

      # Then register all formula tasks via a single data-driven loop (one
      # Croupier::Task.new call site) rather than a literal Task.new block per
      # cell. For large sheets the task table is split across chunk files to
      # keep the Crystal compiler's peak memory low.
      reg_entrypoint, aux_files, did_split = build_task_registration(chunk_prefix, plans)
      source += reg_entrypoint
      source += "\n\n"

      # Finally, run the tasks and print results
      if interactive
        source += generate_tui_mode(initial_values, source_file, intermediate_file)
      else
        source += generate_execution_code(initial_values)
      end

      GeneratedSource.new(source, aux_files, did_split)
    end

    # Write a GeneratedSource to disk: the entrypoint to entrypoint_path, and
    # each aux file into the same directory. The four crystal-build call sites
    # then compile just the entrypoint (which requires its sibling chunks).
    def self.write_generated(gen : GeneratedSource, entrypoint_path : String) : Nil
      File.write(entrypoint_path, gen.entrypoint)
      dir = File.dirname(entrypoint_path)
      gen.aux_files.each do |name, content|
        File.write(File.join(dir, name), content)
      end
    end

    # Generate code to set initial values
    private def generate_setup_code(initial_values : Hash(String, BigFloat | String | Bool), plans : Array(Plan)) : String
      # Collect all unique ranges across the formulas. These come from the
      # same AST walk that produced each plan's dependencies (normalized
      # through CellRefs.parse_range, with the same sheet resolution the
      # fetch calls use), so the emitted initialize_range calls always
      # match the ranges the tasks actually fetch — including sheet names
      # whose escaped literals no text scan could recover.
      ranges = Set(DependencyExtractor::RangeDependency).new
      plans.each do |plan|
        plan.ranges.each { |range| ranges << range }
      end

      # Build a hash of only the cells with actual values
      all_cells = Hash(String, String).new

      # Add user-provided values
      initial_values.each do |key, value|
        value_str = case value
                    when BigFloat then value.to_s
                    when String   then value
                    when Bool     then value.to_s
                    else               ""
                    end
        all_cells[key] = value_str
      end

      # Build setup code
      setup = ""

      # First, initialize all ranges to empty strings (required by Croupier)
      ranges.each do |range|
        bounds = range.bounds
        setup += %(
initialize_range(#{range.sheet.inspect}, #{bounds.start_col.inspect}, #{bounds.start_row}, #{bounds.end_col.inspect}, #{bounds.end_row})
)
      end

      # Then, set the cells with actual values
      unless all_cells.empty?
        setup += %(
# Set initial cell values
initialize_cells(#{all_cells.inspect})

)
      end

      setup
    end

    # Generate code to execute tasks and print results
    private def generate_execution_code(initial_values : Hash(String, BigFloat | String | Bool)) : String
      # Group cells by sheet and organize in grid
      sheets_data = {} of String => Hash(String, Hash(String, String))

      # Add formula cells
      @formulas.each do |key, info|
        parts = key.split("!", 2)
        sheet = parts.size > 1 ? parts[0] : "" # Use empty string for default sheet
        cell = parts.size > 1 ? parts[1] : parts[0]

        sheets_data[sheet] ||= Hash(String, Hash(String, String)).new
        sheets_data[sheet][cell] = {"formula" => info.formula}
      end

      # Add initial value cells (that don't have formulas)
      initial_values.each do |key, _|
        parts = key.split("!", 2)
        sheet = parts.size > 1 ? parts[0] : ""
        cell = parts.size > 1 ? parts[1] : parts[0]

        sheets_data[sheet] ||= Hash(String, Hash(String, String)).new
        # Only add if there's no formula already
        unless sheets_data[sheet][cell]?
          sheets_data[sheet][cell] = {"formula" => ""}
        end
      end

      # Generate code to collect all cell data per sheet. Each sheet's data is
      # built with << appends rather than a single large array literal, so the
      # compiler lowers N cheap append calls instead of one N-element literal.
      sheet_collection_code = sheets_data.map do |sheet, cells|
        sheet_display_name = sheet.empty? ? "(default)" : sheet
        sheet_var_name = sheet.empty? ? "default" : sheet.gsub(/[^a-zA-Z0-9]/, "_")
        sheet_key_prefix = sheet.empty? ? "" : "#{sheet}!"

        appends = cells.map do |cell, data|
          cell_key = "#{sheet_key_prefix}#{cell}"
          %(  sheet_#{sheet_var_name}_data << {cell: #{cell.inspect}, formula: #{data["formula"].inspect}, value: Croupier::TaskManager.get(#{cell_key.inspect}) || "(empty)"})
        end.join("\n")

        "  # Sheet: #{sheet_display_name}
  sheet_#{sheet_var_name}_data = [] of NamedTuple(cell: String, formula: String, value: String)
#{appends}"
      end.join("\n\n")

      sheet_print_code = sheets_data.keys.sort!.map do |sheet|
        sheet_var_name = sheet.empty? ? "default" : sheet.gsub(/[^a-zA-Z0-9]/, "_")
        sheet_display_name = sheet.empty? ? "(default)" : sheet
        "  print_sheet(sheet_#{sheet_var_name}_data, #{sheet_display_name.inspect})"
      end.join("\n")

      # Generate code to display results in a sheet layout
      %(
# Execute all tasks
puts "=== Executing Croupier Tasks ==="
Croupier::TaskManager.run_tasks

# Display results as sheets
puts ""
puts "=== Spreadsheet Results ==="
puts ""

#{sheet_collection_code}

# Helper functions are provided by Sheety::CellRefs (required via
# croupier_helpers) for column letter <-> number conversion.

def print_sheet(data, sheet_name)
  # Find the grid dimensions
  max_col = 0
  max_row = 0

  data.each do |cell|
    if match = cell[:cell].match(/^([A-Z]+)(\\d+)$/)
      col = match[1]
      row = match[2].to_i

      # Convert column to number for comparison
      col_num = Sheety::CellRefs.col_to_num(col)

      max_col = col_num if col_num > max_col
      max_row = row if row > max_row
    end
  end

  # Create a 2D grid
  grid = Array.new(max_row) { Array.new(max_col, "") }

  # Fill the grid
  data.each do |cell|
    if match = cell[:cell].match(/^([A-Z]+)(\\d+)$/)
      col = match[1]
      row = match[2].to_i - 1  # Convert to 0-indexed

      # Convert column to number
      col_num = Sheety::CellRefs.col_to_num(col)

      value = cell[:value]
      formula = cell[:formula]

      # Display: if there's a formula, show it, otherwise just the value
      display = formula.empty? ? value : formula + " -> " + value

      grid[row][col_num - 1] = display
    end
  end

  # Build column headers (A, B, C, ...)
  column_headers = (1..max_col).map { |i| Sheety::CellRefs.num_to_col(i) }

  # Build table data with row numbers
  table_data = (0...max_row).map do |row_idx|
    [row_idx + 1] + grid[row_idx]
  end

  # Create and print the table using Tablo
  table = Tablo::Table.new(table_data) do |t|
    t.add_column("Row", &.[](0).to_s)
    (1..max_col).each do |col_idx|
      t.add_column(column_headers[col_idx - 1], &.[](col_idx).to_s)
    end
  end

  puts table
  puts ""
end

#{sheet_print_code}
puts ""
)
    end

    # Group all parseable formulas by structural shape and emit one
    # `calc_shape_N(...)` helper per unique shape. Populates @shape_assignment
    # so generate_single_task_source can look up each formula's helper index.
    # Formulas with a hard problem (unparseable, unsupported ranges) are
    # skipped here (they keep their inline #VALUE! body) and are absent from
    # @shape_assignment.
    private def generate_shape_helpers(plans : Array(Plan)) : String
      assignment = Hash(String, Int32).new
      # Ordered map: shape key -> {first ast, first sheet, param list, count}.
      shapes = [] of {key: String, ast: AST::Node, sheet: String?, params: Array(CodeGenerator::ReferenceParam)}

      plans.each do |plan|
        ast = plan.ast
        next if ast.nil?

        shape_key = @generator.shape_key(ast)
        existing = shapes.index { |entry| entry[:key] == shape_key }
        if existing
          assignment[plan.info.key] = existing
        else
          params = @generator.reference_params(ast, plan.info.sheet)
          shapes << {key: shape_key, ast: ast, sheet: plan.info.sheet, params: params}
          assignment[plan.info.key] = shapes.size - 1
        end
      end

      @shape_assignment = assignment

      return "" if shapes.empty?

      helpers = String.build do |io|
        io << "# Shared calculation helpers (one per unique formula shape)\n"
        shapes.each_with_index do |entry, index|
          params_decls = entry[:params].each_index.map do |param_index|
            @generator.param_declaration(entry[:params][param_index], param_index)
          end.join(", ")
          # Prefer the param declarations if there are any; otherwise use an
          # empty argument list (a no-arg formula like =42).
          signature = params_decls.empty? ? "" : "(#{params_decls})"
          param_body = @generator.generate_parameterized(entry[:ast], entry[:sheet])

          io << %(# Shape #{index}\n)
          io << %(def calc_shape_#{index}#{signature}\n)
          io << %(  begin\n)
          io << %(    result = (#{param_body})\n)
          io << %(    format_result(result)\n)
          io << %(  rescue e : Exception\n)
          io << %(    "#ERROR: " + (e.message || "Unknown error")\n)
          io << %(  end\n)
          io << %(end\n\n)
        end
      end

      # Trim trailing blank line; caller adds its own spacing.
      helpers.rstrip('\n') + "\n"
    end

    # Build the task-registration section. Returns a triple:
    #   {entrypoint_fragment, aux_files, did_split}
    # For small sheets (<= SPLIT_THRESHOLD formulas) the fragment is the full
    # inline table + loop and aux_files is empty. For large sheets the table is
    # partitioned into chunk files (one TASKS_N = [...] constant each), the
    # fragment requires them and concatenates their constants, and aux_files
    # maps each chunk filename to its content. Either way there is exactly one
    # register_formula_task (and thus one Croupier::Task.new) call site.
    private def build_task_registration(chunk_prefix : String, plans : Array(Plan)) : {String, Hash(String, String), Bool}
      entries = plans.map { |plan| task_entry_source(plan) }

      return {"", Hash(String, String).new, false} if entries.empty?

      loop_body = String.build do |io|
        io << "formula_tasks.each do |task|\n"
        io << "  register_formula_task(task[:id], task[:inputs].call, task[:output], &task[:body])\n"
        io << "end\n"
      end

      # Small sheet: inline the whole table in the entrypoint.
      unless @formulas.size > SPLIT_THRESHOLD
        fragment = String.build do |io|
          io << "# Formula task registration (data-driven loop)\n"
          io << "formula_tasks = [\n"
          entries.each { |entry| io << entry }
          io << "]\n"
          io << loop_body
        end
        return {fragment, Hash(String, String).new, false}
      end

      # Large sheet: split the table across chunk files. Each chunk defines a
      # module-level TASKS_N constant; the entrypoint requires them and
      # concatenates into formula_tasks. Splitting lets the compiler retire
      # per-file structures during semantic analysis, capping peak memory.
      aux_files = Hash(String, String).new
      requires = [] of String
      concats = [] of String

      entries.each_slice(CHUNK_SIZE).with_index do |slice, index|
        filename = "#{chunk_prefix}_tasks_#{index}.cr"
        aux_files[filename] = String.build do |io|
          io << "TASKS_#{index} = [\n"
          slice.each { |entry| io << entry }
          io << "]\n"
        end
        requires << %(require "./#{chunk_prefix}_tasks_#{index}")
        concats << %(formula_tasks.concat(TASKS_#{index}))
      end

      fragment = String.build do |io|
        io << "# Formula task registration (data-driven loop, split across #{aux_files.size} chunk files)\n"
        io << "formula_tasks = [] of NamedTuple(id: String, inputs: Proc(Array(String)), output: String, body: Proc(String))\n"
        requires.each { |require_line| io << require_line << "\n" }
        concats.each { |concat_line| io << concat_line << "\n" }
        io << loop_body
      end
      {fragment, aux_files, true}
    end

    # Build the source for one entry in the formula_tasks table. Mirrors the
    # id/inputs/outputs/body derivation that previously lived in the per-cell
    # literal block, so runtime behavior is unchanged.
    private def task_entry_source(plan : Plan) : String
      info = plan.info
      id = "formula_#{sanitize_key(info.key)}"
      output = "kv://#{info.key}"

      ast = plan.ast
      if ast.nil?
        # Hard problem (unparseable formula, unsupported range): no inputs,
        # #VALUE! body. The problem itself is reported via validation_problems.
        return %(  {id: #{id.inspect}, inputs: ->{ [] of String }, output: #{output.inspect}, body: ->{ "#VALUE!" }},\n)
      end

      dependencies = plan.dependencies

      # Generate the concrete calculation code. Used to derive inputs (by
      # scanning for fetch_cell_range calls) and as the fallback body.
      calc_code = plan.calc_code

      # Build the inputs array expression - use range helpers when ranges appear
      # (both fetch_cell_range and fetch_cell_range_2d share the same argument
      # list, so one pattern catches either helper)
      range_matches = calc_code.scan(/fetch_cell_range(?:_2d)?\(([^)]+)\)/).map(&.[1])

      inputs_expr = if !range_matches.empty?
                      # Multiple ranges - combine them with + operator
                      range_inputs_calls = range_matches.map { |range_params| "range_inputs(#{range_params})" }
                      range_inputs_calls.join(" + ")
                    elsif dependencies.empty?
                      "[] of String"
                    else
                      # Inspect so sheet names that need escaping (a quote, a
                      # backslash) stay valid Crystal literals.
                      "[" + dependencies.map { |dep| %(kv://#{dep}).inspect }.join(", ") + "] of String"
                    end

      # Body: if this formula shares a calc_shape_N helper, call it with the
      # concrete fetch expressions for this cell's reference leaves. Otherwise
      # fall back to the inline calculation.
      body_expr = if (assignment = @shape_assignment) && (index = assignment[info.key]?)
                    call_args = @generator.reference_params(ast, info.sheet).map { |param| @generator.fetch_expression_for(param) }
                    "calc_shape_#{index}(#{call_args.join(", ")})"
                  else
                    %(begin
            result = (#{calc_code})
            format_result(result)
          rescue e : Exception
            "#ERROR: " + (e.message || "Unknown error")
          end)
                  end

      %(  {id: #{id.inspect}, inputs: ->{ #{inputs_expr} }, output: #{output.inspect}, body: ->{ #{body_expr} }},\n)
    end

    # Generate TUI mode
    private def generate_tui_mode(initial_values : Hash(String, BigFloat | String | Bool), source_file : String?, intermediate_file : String?) : String
      # Get the sheet data collection code
      sheets_data = {} of String => Hash(String, Hash(String, String))

      @formulas.each do |key, info|
        parts = key.split("!", 2)
        sheet = parts.size > 1 ? parts[0] : ""
        cell = parts.size > 1 ? parts[1] : parts[0]

        sheets_data[sheet] ||= Hash(String, Hash(String, String)).new
        sheets_data[sheet][cell] = {"formula" => info.formula}
      end

      initial_values.each do |key, _|
        parts = key.split("!", 2)
        sheet = parts.size > 1 ? parts[0] : ""
        cell = parts.size > 1 ? parts[1] : parts[0]

        sheets_data[sheet] ||= Hash(String, Hash(String, String)).new
        unless sheets_data[sheet][cell]?
          sheets_data[sheet][cell] = {"formula" => ""}
        end
      end

      # Generate code to collect all cell data per sheet. Built with << appends
      # rather than a single large array literal, to keep compile-time memory
      # low on big sheets (matches the batch-mode collection form).
      sheet_collection_code = sheets_data.map do |sheet, cells|
        sheet_var_name = sheet.empty? ? "default" : sheet.gsub(/[^a-zA-Z0-9]/, "_")
        sheet_key_prefix = sheet.empty? ? "" : "#{sheet}!"

        appends = cells.map do |cell, data|
          cell_key = "#{sheet_key_prefix}#{cell}"
          %(  sheet_#{sheet_var_name}_data << {cell: #{cell.inspect}, formula: #{data["formula"].inspect}, value: fetch_cell(#{cell_key.inspect})})
        end.join("\n")

        "  sheet_#{sheet_var_name}_data = [] of NamedTuple(cell: String, formula: String, value: String)
#{appends}"
      end.join("\n\n")

      # Build the sheets array and data hash
      sheets_array = sheets_data.keys.sort!.map do |sheet|
        sheet.inspect
      end.join(", ")

      sheets_data_init = sheets_data.keys.sort!.map do |sheet|
        sheet_var_name = sheet.empty? ? "default" : sheet.gsub(/[^a-zA-Z0-9]/, "_")
        "#{sheet.inspect} => sheet_#{sheet_var_name}_data"
      end.join(",\n    ")

      %{
# Execute all tasks
puts "=== Executing Croupier Tasks ==="
Croupier::TaskManager.run_tasks

# Collect cell data for TUI
#{sheet_collection_code}

# Build sheets array and data hash
sheets = [#{sheets_array}]
sheet_data = {
    #{sheets_data_init}
  }

# Create TUI instance
tui = Sheety::TUI.new(sheets, sheet_data) do |sheet, cell_ref, new_value|
  # Update callback: called when user saves a cell edit
  full_key = sheet.empty? ? cell_ref : sheet + "!" + cell_ref

  # Update the value in Croupier store
  Croupier::TaskManager.set(full_key, new_value)

  # Re-run dependent tasks
  Croupier::TaskManager.run_tasks
end

# Set source file for save functionality
#{source_file ? "tui.source_file = #{source_file.inspect}" : ""}

# Set original source file for saves (persists across rebuilds)
#{@original_filename ? "tui.original_source_file = #{@original_filename.inspect}" : (source_file ? "tui.original_source_file = #{source_file.inspect}" : "")}

# Set intermediate file for auto-saves (formula edits)
#{intermediate_file ? "tui.intermediate_file = #{intermediate_file.inspect}" : ""}

# Set value getter callback to fetch fresh values from Croupier store
tui.set_value_getter do |sheet, cell_ref|
  fetch_cell(sheet.empty? ? cell_ref : sheet + "!" + cell_ref)
end

# Set refresh callback to refresh the grid after updates
tui.set_refresh_callback do
  tui.refresh_current_sheet
end

# Restore UI state if available
#{source_file ? %{
# Try to restore cursor position from YAML
begin
  yaml_data = YAML.parse(File.read(#{source_file.inspect}))
  if yaml_data.as_h.has_key?("_ui_state")
    ui_state = yaml_data["_ui_state"]
    if ui_state.as_h.has_key?("active_sheet") && ui_state.as_h.has_key?("active_cell")
      saved_sheet = ui_state["active_sheet"].as_s
      saved_cell = ui_state["active_cell"].as_s
      tui.set_initial_position(saved_sheet, saved_cell)
    end
  end
rescue
  # Ignore errors restoring UI state
end
} : ""}

# Run the TUI
tui.run
}
    end
  end
end
