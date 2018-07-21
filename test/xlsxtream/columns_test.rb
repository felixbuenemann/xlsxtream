require 'test_helper'
require 'xlsxtream/row'

module Xlsxtream
  class ColumnsTest < Minitest::Test
    def test_no_width_column
      column = Columns.new( [ {} ] )
      expected = '<cols><col min="1" max="1"/></cols>'
      actual = column.to_xml
      assert_equal expected, actual
    end

    def test_pixel_width_column
      column = Columns.new( [ { :width_pixels => 2341.5 } ] )
      expected = '<cols><col min="1" max="1" width="2341.5" customWidth="1"/></cols>'
      actual = column.to_xml
      assert_equal expected, actual
    end

    def test_character_width_column

      # https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.column.aspx
      #
      # ...Therefore, if the cell width is 8 characters wide, the value of
      # this attribute must be Truncate([8*7+5]/7*256)/256 = 8.7109375...
      #
      column = Columns.new( [ { :width_chars => 8 } ] )
      expected = '<cols><col min="1" max="1" width="8.7109375" customWidth="1"/></cols>'

      actual = column.to_xml
      assert_equal expected, actual
    end

    def test_mixed_columns
      column = Columns.new( [ {}, { :width_pixels => 61 }, { :width_chars => 14 } ] )
      expected = '<cols><col min="1" max="1"/><col min="2" max="2" width="61" customWidth="1"/><col min="3" max="3" width="14.7109375" customWidth="1"/></cols>'
      actual = column.to_xml
      assert_equal expected, actual
    end
  end
end
