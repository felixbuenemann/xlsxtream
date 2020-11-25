# frozen_string_literal: true
require 'test_helper'
require 'xlsxtream/header_row'

module Xlsxtream
  class HeaderRowTest < Minitest::Test
    def test_header_string_column
      row = HeaderRow.new(['hello'], 1)
      expected = '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_header_symbol_column
      row = HeaderRow.new([:hello], 1)
      expected = '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_header_boolean_column
      row = HeaderRow.new([true], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="b"><v>1</v></c></row>'
      assert_equal expected, actual
      row = HeaderRow.new([false], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="b"><v>0</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_text_boolean_column
      row = HeaderRow.new(['true'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="b"><v>1</v></c></row>'
      assert_equal expected, actual
      row = HeaderRow.new(['false'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="b"><v>0</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_integer_column
      row = HeaderRow.new([1], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_text_integer_column
      row = HeaderRow.new(['1'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_float_column
      row = HeaderRow.new([1.5], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_text_float_column
      row = HeaderRow.new(['1.5'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_date_column
      row = HeaderRow.new([Date.new(1900, 1, 1)], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="4"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_text_date_column
      row = HeaderRow.new(['1900-01-01'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="4"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_invalid_text_date_column
      row = HeaderRow.new(['1900-02-29'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>1900-02-29</t></is></c></row>'
      assert_equal expected, actual
    end

    def test_header_date_time_column
      row = HeaderRow.new([DateTime.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="5"><v>2.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_text_date_time_column
      candidates = [
        '1900-01-01T12:00',
        '1900-01-01T12:00Z',
        '1900-01-01T12:00+00:00',
        '1900-01-01T12:00:00+00:00',
        '1900-01-01T12:00:00.000+00:00',
        '1900-01-01T12:00:00.000000000Z'
      ]
      candidates.each do |timestamp|
        row = HeaderRow.new([timestamp], 1, :auto_format => true)
        actual = row.to_xml
        expected = '<row r="1"><c r="A1" s="5"><v>2.5</v></c></row>'
        assert_equal expected, actual
      end
      row = HeaderRow.new(['1900-01-01T12'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="5"><v>2.5</v></c></row>'
      refute_equal expected, actual
    end

    def test_header_invalid_text_date_time_column
      row = HeaderRow.new(['1900-02-29T12:00'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>1900-02-29T12:00</t></is></c></row>'
      assert_equal expected, actual
    end

    def test_header_time_column
      row = HeaderRow.new([Time.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="5"><v>2.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_header_string_column_with_shared_string_table
      mock_sst = { 'hello' => 0 }
      row = HeaderRow.new(['hello'], 1, :sst => mock_sst)
      expected = '<row r="1"><c r="A1" s="3" t="s"><v>0</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_header_multiple_columns
      row = HeaderRow.new(['foo', nil, 23], 1)
      expected = '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>foo</t></is></c><c r="C1" s="3" t="n"><v>23</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end
  end
end
