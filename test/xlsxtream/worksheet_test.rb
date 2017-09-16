require 'test_helper'
require 'xlsxtream/worksheet'

module Xlsxtream
  class WorksheetTest < Minitest::Test
    def test_empty_worksheet
      io = StringIO.new
      ws = Worksheet.new(io)
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row
      io = StringIO.new
      ws = Worksheet.new(io)
      ws << ['foo']
      ws.add_row ['bar']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row_with_sst_option
      io = StringIO.new
      mock_sst = { 'foo' => 0 }
      ws = Worksheet.new(io, :sst => mock_sst)
      ws << ['foo']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row_with_auto_format_option
      io = StringIO.new
      ws = Worksheet.new(io, :auto_format => true)
      ws << ['1.5']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end
  end
end
