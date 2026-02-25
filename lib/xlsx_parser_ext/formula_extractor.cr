require "compress/zip"
require "xml"

module XlsxParserExt
  # Extracts formulas from Excel worksheet XML files
  # Excel stores formulas in <f> elements within <c> (cell) elements
  # Format: <c r="A1"><f>SUM(B1:B10)</f><v>550</v></c>
  class FormulaExtractor
    # Extract formulas from a specific worksheet in an xlsx file
    # Returns a hash mapping cell references to formula strings
    def self.extract(filename : String, sheet_index : Int32) : Hash(String, String)
      formulas = {} of String => String

      begin
        zip = Compress::Zip::File.new(filename)

        # Worksheet files are named sheet1.xml, sheet2.xml, etc.
        sheet_path = "xl/worksheets/sheet#{sheet_index + 1}.xml"

        if zip[sheet_path]?
          xml_content = zip[sheet_path].open(&.gets_to_end)
          formulas = parse_formulas_from_xml(xml_content)
        end

        zip.close
      rescue ex : Exception
        # If we can't extract formulas, return empty hash
        # This allows the importer to still work with values
        formulas = {} of String => String
      end

      formulas
    end

    # Parse formulas from worksheet XML content
    private def self.parse_formulas_from_xml(xml_content : String) : Hash(String, String)
      formulas = {} of String => String

      begin
        doc = XML.parse(xml_content)

        # Find all <c> (cell) elements using local-name() to handle namespaces
        # Then check for <f> (formula) children
        doc.xpath_nodes("//*[local-name()='c']").each do |cell_node|
          # Get cell reference from r attribute (e.g., "A1")
          cell_ref = cell_node["r"]?

          if cell_ref
            # Look for <f> (formula) element as a child
            formula_nodes = cell_node.xpath_nodes("./*[local-name()='f']")

            if formula_nodes.size > 0
              formula = formula_nodes.first.content.strip

              # Only store non-empty formulas
              unless formula.empty?
                formulas[cell_ref] = formula
              end
            end
          end
        end
      rescue ex : Exception
        # If XML parsing fails, return empty hash
        formulas = {} of String => String
      end

      formulas
    end

    # Extract formulas from all sheets in a workbook
    # Returns array of hashes, one per sheet
    def self.extract_all_sheets(filename : String, sheet_count : Int32) : Array(Hash(String, String))
      all_formulas = Array(Hash(String, String)).new

      sheet_count.times do |i|
        formulas = extract(filename, i)
        all_formulas << formulas
      end

      all_formulas
    end
  end
end
