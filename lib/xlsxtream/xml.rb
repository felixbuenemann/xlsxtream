# frozen_string_literal: true
module Xlsxtream
  module XML
    XML_ESCAPES = {
      '&' => '&amp;',
      '"' => '&quot;',
      '<' => '&lt;',
      '>' => '&gt;',
    }.freeze

    HEX_ESCAPE_REGEXP = /_x[0-9A-Fa-f]{4}_/
    XML_ESCAPE_UNDERSCORE = '_x005f_'

    XML_DECLARATION = %'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'.freeze

    WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/.freeze

    UNSAFE_ATTR_CHARS = /[&"<>]/.freeze
    UNSAFE_VALUE_CHARS = /[&<>]/.freeze

    # http://www.w3.org/TR/REC-xml/#NT-Char:
    # Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    INVALID_XML10_CHARS = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/.freeze

    # ST_Xstring escaping
    ESCAPE_CHAR = lambda { |c| '_x%04X_'.freeze % c.ord }.freeze

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

      # Add underscore to strings that merely look like hex values, preventing manipulation into invalid UTF8
      # Per Microsoft Open Specifications for Excel:
      # Underscore (0x005f): This character shall be escaped only when used to escape the first underscore character in the format _xHHHH_.
      def encode_underscores_using_x005f(string)
        string.gsub(HEX_ESCAPE_REGEXP) do |match|
          match.sub("_", XML_ESCAPE_UNDERSCORE)
        end
      end

      def escape_value(string)
        excel_safe_string = encode_underscores_using_x005f(string)
        excel_safe_string.gsub(UNSAFE_VALUE_CHARS, XML_ESCAPES).gsub(INVALID_XML10_CHARS, &ESCAPE_CHAR)
      end

    end
  end
end
