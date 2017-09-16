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
  end
end
