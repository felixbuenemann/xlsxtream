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
      @has_header_row = options.delete(:has_header_row)

      write_header
    end

    def <<(row)
      options = @has_header_row && @rownum == 1 ? @options.merge(:is_header => true) : @options
      @io << Row.new(row, @rownum, options).to_xml
      @rownum += 1
    end
    alias_method :add_row, :<<

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

      freeze_rows = @options[:freeze_rows].to_i
      freeze_columns = @options[:freeze_columns].to_i
      if freeze_rows > 0 || freeze_columns > 0
        @io << sheet_views_xml(freeze_rows, freeze_columns)
      end

      columns = Array(@options[:columns])
      unless columns.empty?
        @io << Columns.new(columns).to_xml
      end

      @io << XML.strip(<<-XML)
          <sheetData>
      XML
    end

    def sheet_views_xml(rows, columns)
      top_left = String.new('A')
      columns.times { top_left.next! }
      top_left << (rows + 1).to_s

      attributes = String.new
      attributes << %Q{ xSplit="#{columns}"} if columns > 0
      attributes << %Q{ ySplit="#{rows}"} if rows > 0
      active_pane = if rows > 0 && columns > 0 then 'bottomRight'
                    elsif rows > 0 then 'bottomLeft'
                    else 'topRight'
                    end

      %Q{<sheetViews><sheetView workbookViewId="0"><pane#{attributes} topLeftCell="#{top_left}" activePane="#{active_pane}" state="frozen"/></sheetView></sheetViews>}
    end

    def write_footer
      @io << XML.strip(<<-XML)
          </sheetData>
        </worksheet>
      XML
    end
  end
end
