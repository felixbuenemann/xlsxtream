# frozen_string_literal: true
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    def initialize(io, options = {})
      @io = io
      @rownum = 1
      @closed = false
      @options = options
      @xml_merge_cells = nil

      write_header
    end

    def <<(row)
      @io << Row.new(row, @rownum, @options).to_xml
      @rownum += 1
    end
    alias_method :add_row, :<<
    
    # refs: ['A1:B1', 'C1:G1', 'A2:G2', ...]
    def merge_cells(refs)
      xml = String.new(%Q{<mergeCells r="#{refs.length}">})

      refs.each do |ref|
        xml << %Q{<mergeCell ref="#{ref}"/>}
      end

      xml << '</mergeCells>'
      @xml_merge_cells = xml
    end

    def close
      write_footer
      @closed = true
    end

    def closed?
      @closed
    end

    def id
      @options[:id]
    end

    def name
      @options[:name]
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
          #{@xml_merge_cells}
        </worksheet>
      XML
    end
  end
end
