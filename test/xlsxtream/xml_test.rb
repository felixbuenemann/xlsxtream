# frozen_string_literal: true
require 'test_helper'
require 'xlsxtream/xml'

module Xlsxtream
  class XMLTest < Minitest::Test
    def test_header
      expected = %'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
      assert_equal expected, XML.header
    end

    def test_strip
      xml = <<-XML
        <hello id="1">
          <world/>
        </hello>
      XML
      expected = '<hello id="1"><world/></hello>'
      assert_equal expected, XML.strip(xml)
    end

    def test_escape_attr
      unsafe_attribute = '<hello> & "world"'
      expected = '&lt;hello&gt; &amp; &quot;world&quot;'
      assert_equal expected, XML.escape_attr(unsafe_attribute)
    end

    def test_escape_value
      unsafe_value = '<hello> & "world"'
      expected = '&lt;hello&gt; &amp; "world"'
      assert_equal expected, XML.escape_value(unsafe_value)
    end

    def test_escape_value_invalid_xml_chars
      unsafe_value = "The \x07 rings\x00\uFFFE\uFFFF"
      expected = 'The _x0007_ rings_x0000__xFFFE__xFFFF_'
      assert_equal expected, XML.escape_value(unsafe_value)
    end

    def test_escape_value_valid_xml_chars
      safe_value = "\u{10000}\u{10FFFF}"
      expected = safe_value
      assert_equal expected, XML.escape_value(safe_value)
    end

    def test_encode_underscores_using_x005f
      unsafe_value = "_xDcc2_"
      safe_value = "_x005f_xDcc2_"
      assert_equal safe_value, XML.escape_value(unsafe_value)
    end

    def test_encode_underscores_using_x005f_multiple_occurrences
      unsafe_value = "_xDcc2_aa_x3d12_bb_xDea3_cc_xDaa5_"
      safe_value = "_x005f_xDcc2_aa_x005f_x3d12_bb_x005f_xDea3_cc_x005f_xDaa5_"
      assert_equal safe_value, XML.escape_value(unsafe_value)
    end

    def test_not_escaping_regular_underscores
      safe_value = "this_test_does_not_replace_underscores_xDcc2"
      assert_equal safe_value, XML.escape_value(safe_value)
    end
  end
end
