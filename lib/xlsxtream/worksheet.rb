# encoding: utf-8
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    def initialize(io, sst = nil)
      @io = io
      @rownum = 1
      @sst = sst
      write_header
    end

    def <<(row)
      @io << Row.new(row, @rownum, @sst).to_xml
      @rownum += 1
    end
    alias_method :add_row, :<<

    def close
      write_footer
    end

    private

    def write_header
      @io << XML.header
      @io << XML.strip(<<-XML)
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
      XML
    end

    def write_footer
      @io << XML.strip(<<-XML)
          </sheetData>
        </worksheet>
      XML
    end
  end
end
