# frozen_string_literal: true
require "xlsxtream/errors"
require "xlsxtream/xml"
require "xlsxtream/shared_string_table"
require "xlsxtream/worksheet"
require "xlsxtream/io/zip_tricks"
require "xlsxtream/styles/stylesheet"

module Xlsxtream
  class Workbook
    attr_reader :stylesheet, :io

    class << self
      def open(output, options = {})
        workbook = new(output, options)
        if block_given?
          begin
            yield workbook
          ensure
            workbook.close
          end
        else
          workbook
        end
      end
    end

    def initialize(output, options = {})
      if output.nil?
        fail Error, "Xlsxtream::Workbook.new output cannot be nil"
      end
      @options = options
      if options[:io_wrapper]
        fail Deprecation,
          "The Xlsxtream::Workbook.new :io_wrapper option is deprecated. "\
          "Please pass an IO wrapper instance as the first argument instead."
      end

      if output.is_a?(String) || !output.respond_to?(:<<)
        @file = File.open(output, 'wb')
        @io = IO::ZipTricks.new(@file)
      elsif output.respond_to? :add_file
        @file = nil
        @io = output
      else
        @file = nil
        @io = IO::ZipTricks.new(output)
      end

      @sst = SharedStringTable.new
      @worksheets = []

      @stylesheet = Styles::Stylesheet.new(options)
    end

    def add_worksheet(*args, &block)
      if block_given?
        # This method used to be an alias for `write_worksheet`. This was never publicly documented,
        # but to avoid breaking this private API we keep the old behaviour when called with a block.
        Kernel.warn "#{caller.first[/.*:\d+:(?=in `)/]} warning: Calling #{self.class}#add_worksheet with a block is deprecated, use #write_worksheet instead."
        return write_worksheet(*args, &block)
      end

      unless @worksheets.all? { |ws| ws.closed? }
        fail Error, "Close the current worksheet before adding a new one"
      end

      build_worksheet(*args)
    end

    def write_worksheet(*args)
      worksheet = build_worksheet(*args)

      yield worksheet if block_given?
      worksheet.close

      nil
    end

    def close
      write_workbook
      write_styles
      write_sst unless @sst.empty?
      write_workbook_rels
      write_root_rels
      write_content_types
      @io.close if @io.respond_to? :close
      @file.close if @file
      nil
    end

    private
    def build_worksheet(name = nil, options = {})
      if name.is_a? Hash and options.empty?
        options = name
        name = nil
      end

      use_sst = options.fetch(:use_shared_strings, @options[:use_shared_strings])
      auto_format = options.fetch(:auto_format, @options[:auto_format])
      columns = options.fetch(:columns, @options[:columns])
      sst = use_sst ? @sst : nil

      sheet_id = @worksheets.size + 1
      name = name || options[:name] || "Sheet#{sheet_id}"

      @io.add_file "xl/worksheets/sheet#{sheet_id}.xml"

      worksheet = Worksheet.new(self, id: sheet_id, name: name, sst: sst, auto_format: auto_format, columns: columns)
      @worksheets << worksheet

      worksheet
    end

    def write_root_rels
      @io.add_file "_rels/.rels"
      @io << XML.header
      @io << XML.strip(<<-XML)
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def write_workbook
      rid = String.new("rId0")
      @io.add_file "xl/workbook.xml"
      @io << XML.header
      @io << XML.strip(<<-XML)
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <workbookPr date1904="false"/>
          <sheets>
      XML
      @worksheets.each do |worksheet|
        @io << %'<sheet name="#{XML.escape_attr worksheet.name}" sheetId="#{worksheet.id}" r:id="#{rid.next!}"/>'
      end
      @io << XML.strip(<<-XML)
          </sheets>
        </workbook>
      XML
    end

    def write_styles
      @io.add_file "xl/styles.xml"
      @io << @stylesheet.to_xml
    end

    def write_sst
      @io.add_file "xl/sharedStrings.xml"
      @io << XML.header
      @io << %'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{@sst.references}" uniqueCount="#{@sst.size}">'
      @sst.each_key do |string|
        @io << "<si><t>#{XML.escape_value string}</t></si>"
      end
      @io << '</sst>'
    end

    def write_workbook_rels
      rid = String.new("rId0")
      @io.add_file "xl/_rels/workbook.xml.rels"
      @io << XML.header
      @io << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      @worksheets.each do |worksheet|
        @io << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{worksheet.id}.xml"/>'
      end
      @io << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      @io << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' unless @sst.empty?
      @io << '</Relationships>'
    end

    def write_content_types
      @io.add_file "[Content_Types].xml"
      @io << XML.header
      @io << XML.strip(<<-XML)
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      XML
      @io << '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' unless @sst.empty?
      @worksheets.each do |worksheet|
        @io << %'<Override PartName="/xl/worksheets/sheet#{worksheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      end
      @io << '</Types>'
    end
  end
end
