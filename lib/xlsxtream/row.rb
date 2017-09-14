# encoding: utf-8
require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    DATE_STYLE = 1
    TIME_STYLE = 2

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
      @auto_format = options[:auto_format]
    end

    def to_xml
      column = 'A'
      xml = %Q{<row r="#{@rownum}">}

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!

        if @auto_format && value.is_a?(String)
          value = auto_format(value)
        end

        case value
        when Numeric
          xml << %Q{<c r="#{cid}" t="n"><v>#{value}</v></c>}
        when Time, DateTime
          xml << %Q{<c r="#{cid}" s="#{TIME_STYLE}"><v>#{time_to_oa_date(value)}</v></c>}
        when Date
          xml << %Q{<c r="#{cid}" s="#{DATE_STYLE}"><v>#{time_to_oa_date(value)}</v></c>}
        else
          value = value.to_s

          unless value.empty? # no xml output for for empty strings
            value = value.encode(ENCODING) if value.encoding != ENCODING

            if @sst
              xml << %Q{<c r="#{cid}" t="s"><v>#{@sst[value]}</v></c>}
            else
              xml << %Q{<c r="#{cid}" t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>}
            end
          end
        end
      end

      xml << '</row>'
    end

    private

    # Detects and casts numbers, date, time in text
    def auto_format(value)
      case value
      when NUMBER_PATTERN
        value.include?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        Date.parse(value)
      when TIME_PATTERN
        DateTime.parse(value)
      else
        value
      end
    end

    # Converts Time objects to OLE Automation Date
    def time_to_oa_date(time)
      time = time.to_time if time.respond_to?(:to_time)

      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      (time + time.utc_offset).utc.to_f / 24 / 3600 + 25569
    end
  end
end
