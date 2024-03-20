# frozen_string_literal: true
require "date"
require "xlsxtream/core_extension"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    TRUE_STRING = 'true'.freeze
    FALSE_STRING = 'false'.freeze

    DATE_STYLE = 1
    TIME_STYLE = 2

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
      @auto_format = options[:auto_format]
    end

    def to_xml
      column = String.new('A')
      xml = String.new(%Q{<row r="#{@rownum}">})

      @row.each do |value|
        unless value.nil?
          xml << value.to_xslx_value("#{column}#{@rownum}", @auto_format, @sst)
        end
        column.next!
      end

      xml << '</row>'
    end

    # Detects and casts numbers, date, time in text
    class << self
      def auto_format(value)
        case value
        when TRUE_STRING
          true
        when FALSE_STRING
          false
        when NUMBER_PATTERN
          value.include?('.') ? value.to_f : value.to_i
        when DATE_PATTERN
          Date.parse(value) rescue value
        when TIME_PATTERN
          DateTime.parse(value) rescue value
        else
          value
        end
      end
    end
  end
end
