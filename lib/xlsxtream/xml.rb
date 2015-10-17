module Xlsxtream
  module XML
    XML_ESCAPES = {
      '&' => '&amp;',
      '"' => '&quot;',
      '<' => '&lt;',
      '>' => '&gt;',
    }.freeze

    XML_DECLARATION = %'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'.freeze

    WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/.freeze

    UNSAFE_ATTR_CHARS = /[&"<>]/.freeze
    UNSAFE_VALUE_CHARS = /[&<>]/.freeze

    class << self

      def header
        XML_DECLARATION
      end

      def strip(xml)
        xml.gsub(WS_AROUND_TAGS, ''.freeze)
      end

      def escape_attr(string)
        string.gsub(UNSAFE_ATTR_CHARS, XML_ESCAPES)
      end

      def escape_value(string)
        string.gsub(UNSAFE_VALUE_CHARS, XML_ESCAPES)
      end

    end
  end
end
