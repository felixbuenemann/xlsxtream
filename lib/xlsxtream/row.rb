# encoding: utf-8
require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    def initialize(row, rownum, sst = nil)
      @row = row
      @rownum = rownum
      @sst = sst
    end

    def to_xml
      column = 'A'
      xml = %Q{<row r="#{@rownum}">}

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!

        case value
        when Numeric
          xml << %Q{<c r="#{cid}" t="n"><v>#{value}</v></c>}
        when Date, Time, DateTime
          style = value.is_a?(Date) ? 1 : 2
          xml << %Q{<c r="#{cid}" s="#{style}"><v>#{time_to_oa_date(value)}</v></c>}
        else
          value = value.to_s unless value.is_a? String

          if value.empty?
            xml << ''
          else
            value = value.encode(ENCODING) if value.encoding != ENCODING

            if @sst
              xml << %Q{<c r="#{cid}" t="s"><v>#{@sst[value]}</v></c>}
            else
              xml << %Q{<c r="#{cid}" t="inlineStr"><is><t>#{XML.escape_value value}</t></is></c>}
            end
          end
        end
      end

      xml << '</row>'
    end

    private

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
