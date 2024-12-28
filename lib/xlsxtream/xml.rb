# frozen_string_literal: true
module Xlsxtream
  module XML
    XML_ESCAPES = {
      '&' => '&amp;',
      '"' => '&quot;',
      '<' => '&lt;',
      '>' => '&gt;',
    }.freeze

    # Escape first underscore of ST_Xstring sequences in input strings to appear as plaintext in Excel
    HEX_ESCAPE_REGEXP = /_(x[0-9A-Fa-f]{4}_)/
    XML_ESCAPE_UNDERSCORE = '_x005f_\1'

    XML_DECLARATION = %'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

    WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/

    UNSAFE_ATTR_CHARS = /[&"<>]/
    UNSAFE_VALUE_CHARS = /[&<>]/

    # http://www.w3.org/TR/REC-xml/#NT-Char:
    # Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    INVALID_XML10_CHARS = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/

    # ST_Xstring escaping
    ESCAPE_CHAR = lambda { |c| '_x%04X_' % c.ord }

    class << self

      def header
        XML_DECLARATION
      end

      def strip(xml)
        xml.gsub(WS_AROUND_TAGS, '')
      end

      def escape_attr(string)
        string.gsub(UNSAFE_ATTR_CHARS, XML_ESCAPES)
      end

      def escape_value(string)
        string
          .gsub(UNSAFE_VALUE_CHARS, XML_ESCAPES)
          .gsub(HEX_ESCAPE_REGEXP, XML_ESCAPE_UNDERSCORE)
          .gsub(INVALID_XML10_CHARS, &ESCAPE_CHAR)
      end

    end
  end
end
