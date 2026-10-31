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
      @auto_filter = options.delete(:auto_filter)
      @column_count = 0
      @auto_filter_ref = nil

      write_header
    end

    def <<(row)
      options = @has_header_row && @rownum == 1 ? @options.merge(:is_header => true) : @options
      @io << Row.new(row, @rownum, options).to_xml
      @column_count = [@column_count, row.count].max if @auto_filter
      @rownum += 1
    end
    alias_method :add_row, :<<

    def close
      @auto_filter_ref = build_auto_filter_ref
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

    # Cell range covered by the auto filter, e.g. "A1:C5", once the worksheet
    # is closed. nil if the worksheet has no auto filter (or no rows to filter).
    attr_reader :auto_filter_ref

    private

    def build_auto_filter_ref
      return unless @auto_filter && @column_count > 0

      last_column = String.new('A')
      (@column_count - 1).times { last_column.next! }

      "A1:#{last_column}#{@rownum - 1}"
    end

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
      XML
      @io << %Q{<autoFilter ref="#{@auto_filter_ref}"/>} if @auto_filter_ref
      @io << '</worksheet>'
    end
  end
end
