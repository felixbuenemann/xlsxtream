# frozen_string_literal: true
require 'test_helper'
require 'xlsxtream/row'

module Xlsxtream
  class RowTest < Minitest::Test
    DUMMY_IO = StringIO.new
    DUMMY_WORKBOOK = Workbook.new(DUMMY_IO)

    def test_empty_column
      row = Row.new([nil], 1, DUMMY_WORKBOOK)
      expected = '<row r="1"></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_string_column
      row = Row.new(['hello'], 1, DUMMY_WORKBOOK)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_symbol_column
      row = Row.new([:hello], 1, DUMMY_WORKBOOK)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_boolean_column
      row = Row.new([true], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="b"><v>1</v></c></row>'
      assert_equal expected, actual
      row = Row.new([false], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="b"><v>0</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_boolean_column
      row = Row.new(['true'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="b"><v>1</v></c></row>'
      assert_equal expected, actual
      row = Row.new(['false'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="b"><v>0</v></c></row>'
      assert_equal expected, actual
    end

    def test_integer_column
      row = Row.new([1], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_integer_column
      row = Row.new(['1'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_float_column
      row = Row.new([1.5], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_float_column
      row = Row.new(['1.5'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_date_column
      row = Row.new([Date.new(1900, 1, 1)], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="1"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_date_column
      row = Row.new(['1900-01-01'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="1"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_invalid_text_date_column
      row = Row.new(['1900-02-29'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>1900-02-29</t></is></c></row>'
      assert_equal expected, actual
    end

    def test_date_time_column
      row = Row.new([DateTime.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_date_time_column
      candidates = [
        '1900-01-01T12:00',
        '1900-01-01T12:00Z',
        '1900-01-01T12:00+00:00',
        '1900-01-01T12:00:00+00:00',
        '1900-01-01T12:00:00.000+00:00',
        '1900-01-01T12:00:00.000000000Z'
      ]
      candidates.each do |timestamp|
        row = Row.new([timestamp], 1, DUMMY_WORKBOOK, :auto_format => true)
        actual = row.to_xml
        expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
        assert_equal expected, actual
      end
      row = Row.new(['1900-01-01T12'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
      refute_equal expected, actual
    end

    def test_invalid_text_date_time_column
      row = Row.new(['1900-02-29T12:00'], 1, DUMMY_WORKBOOK, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>1900-02-29T12:00</t></is></c></row>'
      assert_equal expected, actual
    end

    def test_time_column
      row = Row.new([Time.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1, DUMMY_WORKBOOK)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_string_column_with_shared_string_table
      mock_sst = { 'hello' => 0 }
      row = Row.new(['hello'], 1, DUMMY_WORKBOOK, :sst => mock_sst)
      expected = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_multiple_columns
      row = Row.new(['foo', nil, 23], 1, DUMMY_WORKBOOK)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c><c r="C1" t="n"><v>23</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_styled_column_with_really_no_style
      row = Row.new([Cell.new('foo')], 1, DUMMY_WORKBOOK)
      expected_style_id = DUMMY_WORKBOOK.stylesheet.default_style_id

      expected = XML.strip(<<-HTML)
        <row r="1">
          <c r="A1" t="inlineStr" s="#{expected_style_id}">
            <is>
              <t>foo</t>
            </is>
          </c>
        </row>
      HTML

      assert_equal expected, row.to_xml
    end

    def test_styled_columns
      cells = [
        Cell.new('Red fill', fill: { color: 'FF0000' }),
        nil,
        Cell.new(23, font: { bold: true, italic: true })
      ]
      row = Row.new(cells, 1, DUMMY_WORKBOOK)

      expected = XML.strip(<<-HTML)
        <row r="1">
          <c r="A1" t="inlineStr" s="3"><is><t>Red fill</t></is></c>
          <c r="C1" t="n" s="4"><v>23</v></c>
        </row>
      HTML

      assert_equal expected, row.to_xml
    end
  end
end
