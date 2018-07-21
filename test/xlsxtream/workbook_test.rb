# frozen_string_literal: true
require 'test_helper'
require 'stringio'
require 'tempfile'
require 'xlsxtream/workbook'
require 'xlsxtream/io/hash'

module Xlsxtream
  class WorksheetTest < Minitest::Test

    def test_workbook_from_path
      tempfile = Tempfile.new('xlsxtream')
      Workbook.open(tempfile.path) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_workbook_from_io
      tempfile = Tempfile.new('xlsxtream')
      Workbook.open(tempfile) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_empty_workbook
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
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
      Workbook.open(iow_spy) do |wb|
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
      Workbook.open(iow_spy) do |wb|
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
      Workbook.new(iow_spy).close
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
      Workbook.new(iow_spy).close
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
      Workbook.open(iow_spy) do |wb|
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
      Workbook.open(iow_spy) do |wb|
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

    def test_write_unnamed_worksheet_with_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(:use_shared_strings => true)
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_worksheet_name_as_option
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |workbook|
        workbook.write_worksheet(name: "foo")
      end
      expected = '<sheet name="foo" sheetId="1" r:id="rId1"/>'
      actual = iow_spy['xl/workbook.xml'][/<sheet [^>]+>/]
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.add_worksheet {}
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData></sheetData></worksheet>'

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options_and_add_rows
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.add_worksheet do |ws|
          ws << ['foo']
          ws.add_row ['bar']
        end
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
        '</sheetData></worksheet>'

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_styles_content
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
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

    def test_custom_font_size
      iow_spy = io_wrapper_spy
      font_options = { :size => 23 }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<sz val="23"/>'
      actual = iow_spy['xl/styles.xml'][/<sz [^>]+>/]
      assert_equal expected, actual
    end

    def test_custom_font_name
      iow_spy = io_wrapper_spy
      font_options = { :name => 'Comic Sans' }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<name val="Comic Sans"/>'
      actual = iow_spy['xl/styles.xml'][/<name [^>]+>/]
      assert_equal expected, actual
    end

    def test_custom_font_family
      iow_spy = io_wrapper_spy
      font_options = { :family => 'Script' }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<family val="4"/>'
      actual = iow_spy['xl/styles.xml'][/<family [^>]+>/]
      assert_equal expected, actual
    end

    def test_font_family_mapping
      tests = {
        nil => 0,
        ''  => 0,
        'ROMAN' => 1,
        :roman => 1,
        'Roman' => 1,
        :swiss => 2,
        :modern => 3,
        :script => 4,
        :decorative => 5
      }
      tests.each do |value, id|
        iow_spy = io_wrapper_spy
        font_options = { :family => value }
        Workbook.open(iow_spy, :font => font_options) {}
        expected = "<family val=\"#{id}\"/>"
        actual = iow_spy['xl/styles.xml'][/<family [^>]+>/]
        assert_equal expected, actual
      end
    end

    def test_invalid_font_family
      iow_spy = io_wrapper_spy
      font_options = { :family => 'Foo' }
      assert_raises Xlsxtream::Error do
        Workbook.open(iow_spy, :font => font_options) {}
      end
    end

    def test_tempfile_is_not_closed
      tempfile = Tempfile.new('workbook')
      Workbook.open(tempfile) {}
      assert_equal false, tempfile.closed?
    ensure
      tempfile && tempfile.close!
    end

    private

    def io_wrapper_spy
      IO::Hash.new(StringIO.new)
    end

  end
end
