# frozen_string_literal: true
require 'test_helper'
require 'stringio'
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

    def test_add_columns_via_worksheet_options
      io = StringIO.new
      ws = Worksheet.new(io, { :columns => [ {}, {}, { :width_pixels => 42 } ] } )
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData></sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_columns_via_worksheet_options_and_add_rows
      io = StringIO.new
      ws = Worksheet.new(io, { :columns => [ {}, {}, { :width_pixels => 42 } ] } )
      ws << ['foo']
      ws.add_row ['bar']
      ws.close
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
      assert_equal expected, io.string
    end

    def test_respond_to_id
      ws = Worksheet.new(StringIO.new, id: 1)
      assert_equal 1, ws.id
    end

    def test_respond_to_name
      ws = Worksheet.new(StringIO.new, name: 'test')
      assert_equal 'test', ws.name
    end
    
    def test_with_merge_cells
      io = StringIO.new
      ws = Worksheet.new(io)
      ws << ['foo']
      ws.add_row ['bar']
      ws.add_row ['foobar']
      ws.merge_cells(['A1:B1', 'A3:B3'])
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
          '<row r="3"><c r="A3" t="inlineStr"><is><t>foobar</t></is></c></row>' \
        '</sheetData>' \
        '<mergeCells r="2">' \
          '<mergeCell ref="A1:B1"/>' \
          '<mergeCell ref="A3:B3"/>' \
        '</mergeCells></worksheet>'
      assert_equal expected, io.string
    end
  end
end
