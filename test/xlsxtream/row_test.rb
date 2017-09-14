require 'test_helper'
require 'xlsxtream/row'

module Xlsxtream
  class RowTest < Minitest::Test
    def test_empty_column
      row = Row.new([nil], 1)
      expected = '<row r="1"></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_string_column
      row = Row.new(['hello'], 1)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_symbol_column
      row = Row.new([:hello], 1)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_integer_column
      row = Row.new([1], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_integer_column
      row = Row.new(['1'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'
      assert_equal expected, actual
    end

    def test_float_column
      row = Row.new([1.5], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_float_column
      row = Row.new(['1.5'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_date_column
      row = Row.new([Date.new(1900, 1, 1)], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="1"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_text_date_column
      row = Row.new(['1900-01-01'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="1"><v>2.0</v></c></row>'
      assert_equal expected, actual
    end

    def test_date_time_column
      row = Row.new([DateTime.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1)
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
        row = Row.new([timestamp], 1, :auto_format => true)
        actual = row.to_xml
        expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
        assert_equal expected, actual
      end
      row = Row.new(['1900-01-01T12'], 1, :auto_format => true)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
      refute_equal expected, actual
    end

    def test_time_column
      row = Row.new([Time.new(1900, 1, 1, 12, 0, 0, '+00:00')], 1)
      actual = row.to_xml
      expected = '<row r="1"><c r="A1" s="2"><v>2.5</v></c></row>'
      assert_equal expected, actual
    end

    def test_string_column_with_shared_string_table
      mock_sst = { 'hello' => 0 }
      row = Row.new(['hello'], 1, :sst => mock_sst)
      expected = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end

    def test_multiple_columns
      row = Row.new(['foo', nil, 23], 1)
      expected = '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c><c r="C1" t="n"><v>23</v></c></row>'
      actual = row.to_xml
      assert_equal expected, actual
    end
  end
end
