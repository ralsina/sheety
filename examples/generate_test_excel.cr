#!/usr/bin/env crystal
# Helper script to generate a test Excel file
# This creates the minimum XML structure required for an .xlsx file

require "file_utils"

# Create a minimal test Excel file
def create_test_excel(filename)
  File.delete(filename) if File.exists?(filename)

  # Use shell commands to create the zip
  temp_dir = "/tmp/xlsx_temp_#{Process.pid}"
  Dir.mkdir(temp_dir)

  # Create directory structure
  Dir.mkdir("#{temp_dir}/_rels")
  Dir.mkdir("#{temp_dir}/xl")
  Dir.mkdir("#{temp_dir}/xl/_rels")
  Dir.mkdir("#{temp_dir}/xl/worksheets")
  Dir.mkdir("#{temp_dir}/xl/theme")
  Dir.mkdir("#{temp_dir}/docProps")

  # [Content_Types].xml
  content_types = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  </Types>
  XML
  File.write("#{temp_dir}/[Content_Types].xml", content_types)

  # _rels/.rels
  rels = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>
  XML
  File.write("#{temp_dir}/_rels/.rels", rels)

  # xl/workbook.xml
  workbook = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
      <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
      <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    </sheets>
  </workbook>
  XML
  File.write("#{temp_dir}/xl/workbook.xml", workbook)

  # xl/_rels/workbook.xml.rels
  workbook_rels = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  </Relationships>
  XML
  File.write("#{temp_dir}/xl/_rels/workbook.xml.rels", workbook_rels)

  # xl/worksheets/sheet1.xml
  sheet1 = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:D4"/>
    <sheetData>
      <row r="1">
        <c r="A1" t="n"><v>100</v></c>
        <c r="B1" t="s"><v>0</v></c>
        <c r="C1" t="n"><v>5</v></c>
        <c r="D1"><f>SUM(A1:A20)</f><v>0</v></c>
      </row>
      <row r="2">
        <c r="A2" t="n"><v>200</v></c>
        <c r="B2" t="s"><v>1</v></c>
        <c r="C2" t="n"><v>3</v></c>
        <c r="D2"><f>MAX(A1:A20)</f><v>0</v></c>
      </row>
      <row r="3">
        <c r="A3" t="n"><v>300</v></c>
        <c r="B3"><f>CONCAT(B1," ",B2)</f><v>0</v></c>
        <c r="C3"><f>IF(C1>C2,"Yes","No")</f><v>0</v></c>
        <c r="D3"><f>MIN(A1:A20)</f><v>0</v></c>
      </row>
      <row r="4">
        <c r="A4" t="n"><v>150</v></c>
        <c r="D4"><f>AVERAGE(A1:A20)</f><v>0</v></c>
      </row>
    </sheetData>
  </worksheet>
  XML
  File.write("#{temp_dir}/xl/worksheets/sheet1.xml", sheet1)

  # xl/worksheets/sheet2.xml
  sheet2 = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:A3"/>
    <sheetData>
      <row r="1">
        <c r="A1" t="n"><v>100</v></c>
      </row>
      <row r="2">
        <c r="A2"><f>Sheet1!D3*2</f><v>0</v></c>
      </row>
      <row r="3">
        <c r="A3"><f>SUM(Sheet1!A1:A2)</f><v>0</v></c>
      </row>
    </sheetData>
  </worksheet>
  XML
  File.write("#{temp_dir}/xl/worksheets/sheet2.xml", sheet2)

  # xl/theme/theme1.xml (minimal)
  theme = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
      <a:colorScheme name="Office">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
        <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
        <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
        <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
        <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
        <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
        <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
        <a:accent6><a:srgbClr val="F79646"/></a:accent6>
        <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
        <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
      </a:colorScheme>
    </a:themeElements>
  </a:theme>
  XML
  File.write("#{temp_dir}/xl/theme/theme1.xml", theme)

  # xl/styles.xml (minimal)
  styles = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="0"/>
    <fonts count="1">
      <font><sz val="11"/><name val="Calibri"/></font>
    </fonts>
    <fills count="1">
      <fill><patternFill patternType="none"/></fill>
    </fills>
    <borders count="1">
      <border><left/><right/><top/><bottom/></border>
    </borders>
    <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellXfs>
  </styleSheet>
  XML
  File.write("#{temp_dir}/xl/styles.xml", styles)

  # xl/sharedStrings.xml (for our string values)
  shared_strings = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    <si><t>Hello</t></si>
    <si><t>World</t></si>
  </sst>
  XML
  File.write("#{temp_dir}/xl/sharedStrings.xml", shared_strings)

  # docProps/core.xml
  core = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Sheety Test Generator</dc:creator>
    <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
  </cp:coreProperties>
  XML
  File.write("#{temp_dir}/docProps/core.xml", core)

  # docProps/app.xml
  app = <<-XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Sheety Test Generator</Application>
  </Properties>
  XML
  File.write("#{temp_dir}/docProps/app.xml", app)

  # Move to the final location and zip
  final_dir = File.dirname(File.expand_path(filename))
  final_name = File.basename(filename)

  # Create the zip file using system command
  Process.run("zip", ["-r", "-q", File.join(final_dir, final_name), "."], chdir: temp_dir, output: :inherit, error: :inherit)

  # Cleanup
  FileUtils.rm_r(temp_dir)

  puts "Test Excel file created: #{filename}"
end

# Create the test file
create_test_excel("examples/test_excel.xlsx")
