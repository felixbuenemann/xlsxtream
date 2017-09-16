require 'test_helper'
require 'xlsxtream/workbook'
require 'xlsxtream/io/hash'

module Xlsxtream
  class WorksheetTest < Minitest::Test

    def test_empty_workbook
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) {}
      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets></sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sheet
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) do |wb|
        wb.add_worksheet
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sst
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) do |wb|
        wb.add_worksheet(nil, use_shared_strings: true) do |ws|
          ws << ['foo']
        end
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData>' \
              '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' \
            '</sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/sharedStrings.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' \
            '<si><t>foo</t></si>' \
          '</sst>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_root_relations
      iow_spy = io_wrapper_spy
      Workbook.new(nil, :io_wrapper => iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
        '</Relationships>'
      actual = iow_spy['_rels/.rels']
      assert_equal expected, actual
    end

    def test_content_types
      iow_spy = io_wrapper_spy
      Workbook.new(nil, :io_wrapper => iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
          '<Default Extension="xml" ContentType="application/xml"/>' \
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' \
        '</Types>'
      actual = iow_spy['[Content_Types].xml']
      assert_equal expected, actual
    end

    def test_write_multiple_worksheets
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) do |wb|
        wb.write_worksheet
        wb.write_worksheet
      end

      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
              '<sheet name="Sheet2" sheetId="2" r:id="rId2"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>',
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>',
        'xl/worksheets/sheet2.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_write_named_worksheet
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) do |wb|
        wb.write_worksheet('foo')
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="foo" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_styles_content
      iow_spy = io_wrapper_spy
      Workbook.open(nil, :io_wrapper => iow_spy) {}
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
          '<numFmts count="2">' \
            '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>' \
            '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>' \
          '</numFmts>' \
          '<fonts count="1">' \
            '<font>' \
              '<sz val="12"/>' \
              '<name val="Calibri"/>' \
              '<family val="2"/>' \
            '</font>' \
          '</fonts>' \
          '<fills count="2">' \
            '<fill>' \
              '<patternFill patternType="none"/>' \
            '</fill>' \
            '<fill>' \
              '<patternFill patternType="gray125"/>' \
            '</fill>' \
          '</fills>' \
          '<borders count="1">' \
            '<border/>' \
          '</borders>' \
          '<cellStyleXfs count="1">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' \
          '</cellStyleXfs>' \
          '<cellXfs count="3">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' \
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
          '</cellXfs>' \
          '<cellStyles count="1">' \
            '<cellStyle name="Normal" xfId="0" builtinId="0"/>' \
          '</cellStyles>' \
          '<dxfs count="0"/>' \
          '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' \
        '</styleSheet>'
      actual = iow_spy['xl/styles.xml']
      assert_equal expected, actual
    end

    private

    def io_wrapper_spy
      Class.new do
        @iow = IO::Hash.new(StringIO.new)
        class << self
          def new(*)
            @iow
          end

          def [](path)
            @iow[path]
          end

          def to_h
            @iow.to_h
          end
        end
      end
    end

  end
end
