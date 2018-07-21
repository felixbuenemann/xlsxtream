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
  end
end
