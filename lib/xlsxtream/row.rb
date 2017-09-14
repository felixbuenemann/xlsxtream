# encoding: utf-8
require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze # yyyy-mm-dd
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}/.freeze # yyyy-mm-ddThh:mm:ss

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

        if value.is_a?(Numeric) || number_string?(value)
          value = value.include?('.') ? value.to_f : value.to_i if number_string?(value)
          xml << %Q{<c r="#{cid}" t="n"><v>#{value}</v></c>}

        elsif value.is_a?(Time) || value.is_a?(DateTime) || time_string?(value)
          value = DateTime.parse(value) if time_string?(value)
          xml << %Q{<c r="#{cid}" s="#{TIME_STYLE}"><v>#{time_to_oa_date(value)}</v></c>}

        elsif value.is_a?(Date) || date_string?(value)
          value = Date.parse(value) if date_string?(value)
          xml << %Q{<c r="#{cid}" s="#{DATE_STYLE}"><v>#{time_to_oa_date(value)}</v></c>}

        else
          value = value.to_s unless value.is_a?(String)

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
    def number_string?(value)
      @auto_format && value.is_a?(String) && value =~ NUMBER_PATTERN
    end

    def date_string?(value)
      @auto_format && value.is_a?(String) && value =~ DATE_PATTERN
    end

    def time_string?(value)
      @auto_format && value.is_a?(String) && value =~ TIME_PATTERN
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
