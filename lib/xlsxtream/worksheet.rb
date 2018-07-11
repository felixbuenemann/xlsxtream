# encoding: utf-8
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    def initialize(io, options = {})
      @io = io
      @rownum = 1
      @options = options
      @sheetdata_written = false

      write_header
    end

    # If you want to specify custom column widths, do so using this
    # method. See Xlsstream::Columns#initialize for parameter details.
    # This MUST be called before #add_row or #<<.
    #
    def add_columns(column_options_array)
      @io << Columns.new(column_options_array).to_xml
    end

    def <<(row)
      unless @sheetdata_written
        @sheetdata_written = true
        @io << '<sheetData>'
      end

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
    end

    def write_footer
      unless @sheetdata_written
        @sheetdata_written = true
        @io << '<sheetData>'
      end

      @io << XML.strip(<<-XML)
          </sheetData>
        </worksheet>
      XML
    end
  end
end
