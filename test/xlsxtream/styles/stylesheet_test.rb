require 'test_helper'

module Xlsxtream; module Styles
  class StylesheetTest < Minitest::Test
    DEFAULT_NUM_FORMAT_ID = 0
    DATE_NUM_FORMAT_ID = 164
    TIME_NUM_FORMAT_ID = 165

    SUPPORTED_OPTIONS = {
      fill: { color: 'FF0000' }, font: { bold: true, italic: true }
    }

    UNSUPPORTED_OPTIONS = {
      foo: { color: 'FF0000' }, bar: { bold: true, italic: true }
    }

    def test_constants
      assert_equal DEFAULT_NUM_FORMAT_ID, Stylesheet::DEFAULT_NUM_FORMAT_ID
      assert_equal DATE_NUM_FORMAT_ID, Stylesheet::DATE_NUM_FORMAT_ID
      assert_equal TIME_NUM_FORMAT_ID, Stylesheet::TIME_NUM_FORMAT_ID
    end

    def test_initialize_populates_basic_xfs
      workbook = Workbook.new(StringIO.new)
      assert_equal 3, peek_num_of_styles(workbook.stylesheet)
    end

    def test_initialize_respects_global_options
      workbook = Workbook.new(StringIO.new, SUPPORTED_OPTIONS)

      # does not create a new default
      assert_equal 3, peek_num_of_styles(workbook.stylesheet)

      # the default style is set using global options
      actual_style_id = workbook.stylesheet.style_id(Cell.new('foo', SUPPORTED_OPTIONS))
      assert_equal workbook.stylesheet.default_style_id, actual_style_id
    end

    def test_style_id
      workbook = Workbook.new(StringIO.new)
      initial_num_styles = peek_num_of_styles(workbook.stylesheet)

      # creates a new style for a cell with the new style
      cell = Cell.new('foo', SUPPORTED_OPTIONS)
      assert_equal initial_num_styles, workbook.stylesheet.style_id(cell)

      # does not create a new style for the cell with the same style
      cell = Cell.new('bar', SUPPORTED_OPTIONS.dup)
      assert_equal initial_num_styles, workbook.stylesheet.style_id(cell)

      # does not respect unsupported options
      cell = Cell.new('foo', UNSUPPORTED_OPTIONS)
      assert_equal 0, workbook.stylesheet.style_id(cell)
    end

    def test_to_xml
      workbook = Workbook.new(StringIO.new)

      cell = Cell.new('foo', SUPPORTED_OPTIONS)
      workbook.stylesheet.style_id(cell)

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
          '<numFmts count="2">' \
            '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>' \
            '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>' \
          '</numFmts>' \
          '<fonts count="2">' \
            '<font>' \
              '<sz val="12"/>' \
              '<name val="Calibri"/>' \
              '<family val="2"/>' \
            '</font>' \
            '<font>' \
              '<sz val="12"/>' \
              '<name val="Calibri"/>' \
              '<family val="2"/>' \
              '<b val="true"/>' \
              '<i val="true"/>' \
            '</font>' \
          '</fonts>' \
          '<fills count="2">' \
            '<fill>' \
            '</fill>' \
            '<fill>' \
              '<patternFill patternType="solid">' \
                '<fgColor rgb="FF0000"/>' \
              '</patternFill>' \
            '</fill>' \
          '</fills>' \
          '<borders count="1">' \
            '<border/>' \
          '</borders>' \
          '<cellStyleXfs count="1">' \
            '<xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/>' \
          '</cellStyleXfs>' \
          '<cellXfs count="4">' \
            '<xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/>' \
            '<xf xfId="0" numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' \
            '<xf xfId="0" numFmtId="165" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' \
            '<xf xfId="0" numFmtId="0" borderId="0" fontId="1" fillId="1" applyFill="1"/>' \
          '</cellXfs>' \
          '<cellStyles count="1">' \
            '<cellStyle name="Normal" xfId="0" builtinId="0"/>' \
          '</cellStyles>' \
          '<dxfs count="0"/>' \
          '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' \
        '</styleSheet>'

      actual = workbook.stylesheet.to_xml
      assert_equal expected, actual
    end

    private

    def peek_num_of_styles(stylesheet)
      stylesheet.instance_variable_get("@xfs").size
    end
  end
end; end
