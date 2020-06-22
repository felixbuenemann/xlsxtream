# frozen_string_literal: true
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    attr_reader :id, :name

    def initialize(io, id, name, options = {})
      @io = io
      @id = id
      @name = name
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

      columns = Array(@options[:columns])
      unless columns.empty?
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
