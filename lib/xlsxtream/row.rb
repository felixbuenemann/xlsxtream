# encoding: utf-8
require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row
    def initialize(row, rownum, sst = nil)
      @row = row
      @rownum = rownum
      @sst = sst
      @encoding = Encoding.find("UTF-8")
    end

    def to_xml
      column = 'A'
      @row.reduce(%'<row r="#@rownum">') do |xml, value|
        cid = "#{column}#@rownum"
        column.next!
        xml << case value
        when Fixnum, Float
          %'<c r="#{cid}" t="n"><v>#{value}</v></c>'
        when Date, Time, DateTime
          style = value.is_a?(Date) ? 1 : 2
          %'<c r="#{cid}" s="#{style}"><v>#{time_to_oa_date value}</v></c>'
        else
          value = value.to_s unless value.is_a? String
          if value.empty?
            ''
          else
            value = value.encode(@encoding) if value.encoding != @encoding
            if @sst
              %'<c r="#{cid}" t="s"><v>#{@sst[value]}</v></c>'
            else
              %'<c r="#{cid}" t="inlineStr"><is><t>#{XML.escape_value value}</t></is></c>'
            end
          end
        end
      end << '</row>'
    end

    private

    # Converts Time objects to OLE Automation Date
    def time_to_oa_date(time)
      time = time.respond_to?(:to_time) ? time.to_time : time
      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      (time + time.utc_offset).utc.to_f / 24 / 3600 + 25569
    end
  end
end
