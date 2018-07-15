# encoding: utf-8
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    def initialize(io, options = {})
      @io = io
      @rownum = 1
      @options = options

      write_header
    end

    def <<(row)
      @io << Row.new(row, @rownum, @options).to_xml
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
      XML

      columns = @options[:columns]
      if columns.is_a?(Array)
        @io << Columns.new(columns).to_xml
      end

      @io << XML.strip(<<-XML)
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
