require "compress/zip"
require "xml"
require "../cell_refs"

module Sheety
  class FormulaExtractor
    # Extract formulas from a specific worksheet in an xlsx file
    # Returns a hash mapping cell references to formula strings
    def self.extract(filename : String, sheet_index : Int32) : Hash(String, String)
      formulas = {} of String => String

      begin
        Compress::Zip::File.open(filename) do |zip|
          # Resolve actual worksheet path from workbook relationships
          sheet_path = resolve_sheet_path(zip, sheet_index)

          if sheet_path && zip[sheet_path]?
            xml_content = zip[sheet_path].open(&.gets_to_end)
            formulas = parse_formulas_from_xml(xml_content)
          end
        end
      rescue Exception
        # If we can't extract formulas, return empty hash
        # This allows the importer to still work with values
        formulas = {} of String => String
      end

      formulas
    end

    # Extract the XML cell type ("n"/absent = numeric, "s" = shared string,
    # "b" = boolean, ...) per cell reference. Needed because xlsx-parser
    # returns integer-valued numeric cells as strings, and only this
    # attribute knows whether a "100" was a number or text.
    def self.extract_cell_types(filename : String, sheet_index : Int32) : Hash(String, String)
      types = {} of String => String

      begin
        Compress::Zip::File.open(filename) do |zip|
          sheet_path = resolve_sheet_path(zip, sheet_index)

          if sheet_path && zip[sheet_path]?
            doc = XML.parse(zip[sheet_path].open(&.gets_to_end))
            doc.xpath_nodes("//*[local-name()='c']").each do |cell_node|
              if ref = cell_node["r"]?
                types[ref] = cell_node["t"]? || "n"
              end
            end
          end
        end
      rescue Exception
        # Without type info the importer keeps whatever xlsx-parser produced.
        types = {} of String => String
      end

      types
    end

    # Resolves the actual worksheet XML path from workbook relationships.
    # This is necessary because worksheet files may not be named sheet1.xml, sheet2.xml, etc.
    private def self.resolve_sheet_path(zip : Compress::Zip::File, sheet_index : Int32) : String?
      # Parse workbook.xml to get sheet IDs
      workbook = XML.parse(zip["xl/workbook.xml"].open(&.gets_to_end))
      sheets_nodes = workbook.xpath_nodes("//*[name()='sheet']")

      return if sheet_index >= sheets_nodes.size

      sheet_node = sheets_nodes[sheet_index]
      sheet_id = sheet_node["id"]?

      return unless sheet_id

      # Parse workbook relationships to find the actual worksheet file
      rels = XML.parse(zip["xl/_rels/workbook.xml.rels"].open(&.gets_to_end))
      sheet_file = rels.xpath_string(
        "string(//*[name()='Relationship' and contains(@Id,'#{sheet_id}')]/@Target)"
      )

      # Target is relative to xl/ directory
      sheet_file.empty? ? nil : "xl/#{sheet_file}"
    rescue Exception
      # Fallback to simple naming if relationship parsing fails
      "xl/worksheets/sheet#{sheet_index + 1}.xml"
    end

    # Parse formulas from worksheet XML content.
    #
    # Excel stores a formula dragged across many cells as one shared master
    # (with content) plus per-cell slaves that only carry the shared index,
    # so slaves are resolved by translating the master's references by the
    # offset between the two cells.
    private def self.parse_formulas_from_xml(xml_content : String) : Hash(String, String)
      formulas = {} of String => String

      begin
        doc = XML.parse(xml_content)
        cell_nodes = doc.xpath_nodes("//*[local-name()='c']")

        # First pass: collect shared-formula masters ("si" => {formula, cell}).
        # The formula type attribute is `t` per the OOXML spec; some producers
        # use `type`, so accept either.
        shared_masters = Hash(String, Tuple(String, String)).new
        cell_nodes.each do |cell_node|
          cell_ref = cell_node["r"]?
          next unless cell_ref

          formula_nodes = cell_node.xpath_nodes("./*[local-name()='f']")
          next if formula_nodes.empty?
          formula_node = formula_nodes.first
          next unless formula_node["t"]? == "shared" || formula_node["type"]? == "shared"

          shared_index = formula_node["si"]?
          next unless shared_index

          content = formula_node.content.strip
          shared_masters[shared_index] = {content, cell_ref} unless content.empty?
        end

        # Second pass: resolve every formula cell.
        cell_nodes.each do |cell_node|
          # Get cell reference from r attribute (e.g., "A1")
          cell_ref = cell_node["r"]?
          next unless cell_ref

          # Look for <f> (formula) element as a child
          formula_nodes = cell_node.xpath_nodes("./*[local-name()='f']")
          next if formula_nodes.empty?
          formula_node = formula_nodes.first

          formula_type = formula_node["t"]? || formula_node["type"]?
          content = formula_node.content.strip

          formula = if formula_type == "shared"
                      shared_index = formula_node["si"]?
                      if !content.empty?
                        content
                      elsif shared_index && (master = shared_masters[shared_index]?)
                        translate_shared(master[0], master[1], cell_ref)
                      else
                        # Shared formula whose master is missing: keep a placeholder so
                        # the cell still surfaces (and the generator warns about it).
                        "SHARED_FORMULA(#{shared_index})"
                      end
                    else
                      content
                    end

          # Only store non-empty formulas
          unless formula.empty?
            formulas[cell_ref] = formula
          end
        end
      rescue Exception
        # If XML parsing fails, return empty hash
        formulas = {} of String => String
      end

      formulas
    end

    # Translate a shared formula from its master cell to a slave cell by
    # offsetting relative references (Excel semantics). $-anchored row/column
    # parts stay put; text inside string literals is untouched.
    #
    # This is a pragmatic translation: function names followed by "(" (LOG10),
    # sheet names (Sheet1, and the "A1!" edge), and named ranges don't match
    # the cell-reference token pattern. References shifted off the sheet
    # become #REF!, as in Excel.
    def self.translate_shared(master : String, from_ref : String, to_ref : String) : String
      from = CellRefs.parse_ref(from_ref)
      to = CellRefs.parse_ref(to_ref)
      return master if from.nil? || to.nil?

      dcol = to[:col] - from[:col]
      drow = to[:row] - from[:row]
      return master if dcol == 0 && drow == 0

      # A relative cell reference: optional $, 1-3 letters, optional $,
      # digits. The trailing lookahead keeps function names (LOG10() and
      # sheet-name tokens (A1!) from matching.
      cell_token = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Z0-9_(!])/i

      String.build do |io|
        in_string = false
        position = 0
        while position < master.size
          char = master[position]
          if char == '"'
            in_string = !in_string
            io << char
            position += 1
            next
          end
          if in_string
            io << char
            position += 1
            next
          end

          # A reference must start on a token boundary.
          boundary_ok = position.zero? || !(master[position - 1].alphanumeric? || master[position - 1] == '_' || master[position - 1] == '$')
          match = boundary_ok ? master.match(cell_token, position) : nil

          # #match(str, pos) can return a match starting after pos; only a
          # match starting exactly here is a reference token.
          if match && match.begin(0) == position
            anchor_col = match[1] == "$"
            anchor_row = match[3] == "$"
            col = CellRefs.col_to_num(match[2])
            row = match[4].to_i

            new_col = anchor_col ? col : col + dcol
            new_row = anchor_row ? row : row + drow

            if new_col < 1 || new_row < 1
              io << "#REF!"
            else
              io << (anchor_col ? "$" : "")
              io << CellRefs.num_to_col(new_col)
              io << (anchor_row ? "$" : "")
              io << new_row
            end
            position = match.end(0)
          else
            io << char
            position += 1
          end
        end
      end
    end
  end
end
