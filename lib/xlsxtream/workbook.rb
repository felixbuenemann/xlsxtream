# frozen_string_literal: true
require "xlsxtream/errors"
require "xlsxtream/xml"
require "xlsxtream/shared_string_table"
require "xlsxtream/worksheet"
require "xlsxtream/zip_kit_writer"

module Xlsxtream
  class Workbook

    FONT_FAMILY_IDS = {
      ''           => 0,
      'roman'      => 1,
      'swiss'      => 2,
      'modern'     => 3,
      'script'     => 4,
      'decorative' => 5
    }.freeze

    DEFAULT_DATE_FORMAT = 'yyyy\\-mm\\-dd'
    DEFAULT_DATETIME_FORMAT = 'yyyy\\-mm\\-dd hh:mm:ss'

    # Built-in formats which Excel renders using the locale of the viewing machine
    LOCALE_DATE_FORMAT_ID = 14
    LOCALE_DATETIME_FORMAT_ID = 22

    FIRST_CUSTOM_FORMAT_ID = 164

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
      @options = options
      @custom_num_fmts = {}
      @date_format_id = resolve_num_fmt(:date_format, DEFAULT_DATE_FORMAT, LOCALE_DATE_FORMAT_ID)
      @datetime_format_id = resolve_num_fmt(:datetime_format, DEFAULT_DATETIME_FORMAT, LOCALE_DATETIME_FORMAT_ID)
      @writer = ZipKitWriter.with_output_to(output)
      @sst = SharedStringTable.new
      @worksheets = []
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
      @writer.close

      nil
    end

    private

    def resolve_num_fmt(option_name, default_format_code, locale_format_id)
      format = @options.fetch(option_name, nil) || default_format_code
      case format
      when :locale
        locale_format_id
      when String
        @custom_num_fmts[format] ||= FIRST_CUSTOM_FORMAT_ID + @custom_num_fmts.size
      else
        fail Error, "Invalid #{option_name} #{format.inspect}, must be :locale or a format code String"
      end
    end

    def build_worksheet(name = nil, options = {})
      if name.is_a? Hash and options.empty?
        options = name
        name = nil
      end

      use_sst = options.fetch(:use_shared_strings, @options[:use_shared_strings])
      auto_format = options.fetch(:auto_format, @options[:auto_format])
      columns = options.fetch(:columns, @options[:columns])
      has_header_row = options.fetch(:has_header_row, @options[:has_header_row])
      sst = use_sst ? @sst : nil

      sheet_id = @worksheets.size + 1
      name = name || options[:name] || "Sheet#{sheet_id}"

      @writer.add_file "xl/worksheets/sheet#{sheet_id}.xml"

      worksheet = Worksheet.new(@writer, :id => sheet_id, :name => name, :sst => sst, :auto_format => auto_format, :columns => columns, :has_header_row => has_header_row)
      @worksheets << worksheet

      worksheet
    end

    def write_root_rels
      @writer.add_file "_rels/.rels"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def write_workbook
      rid = String.new("rId0")
      @writer.add_file "xl/workbook.xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <workbookPr date1904="false"/>
          <sheets>
      XML
      @worksheets.each do |worksheet|
        @writer << %'<sheet name="#{XML.escape_attr worksheet.name}" sheetId="#{worksheet.id}" r:id="#{rid.next!}"/>'
      end
      @writer << XML.strip(<<-XML)
          </sheets>
        </workbook>
      XML
    end

    def write_styles
      font_options = @options.fetch(:font, {})
      font_size = font_options.fetch(:size, 12).to_s
      font_name = font_options.fetch(:name, 'Calibri').to_s
      font_family = font_options.fetch(:family, 'Swiss').to_s.downcase
      font_family_id = FONT_FAMILY_IDS[font_family] or fail Error,
        "Invalid font family #{font_family}, must be one of "\
        + FONT_FAMILY_IDS.keys.map(&:inspect).join(', ')

      @writer.add_file "xl/styles.xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      XML
      unless @custom_num_fmts.empty?
        @writer << %'<numFmts count="#{@custom_num_fmts.size}">'
        @custom_num_fmts.each do |format_code, id|
          @writer << %'<numFmt numFmtId="#{id}" formatCode="#{XML.escape_attr format_code}"/>'
        end
        @writer << '</numFmts>'
      end
      @writer << XML.strip(<<-XML)
          <fonts count="2">
            <font>
              <sz val="#{XML.escape_attr font_size}"/>
              <name val="#{XML.escape_attr font_name}"/>
              <family val="#{font_family_id}"/>
            </font>
            <font>
              <b val="1"/>
              <sz val="#{XML.escape_attr font_size}"/>
              <name val="#{XML.escape_attr font_name}"/>
              <family val="#{font_family_id}"/>
            </font>
          </fonts>
          <fills count="2">
            <fill>
              <patternFill patternType="none"/>
            </fill>
            <fill>
              <patternFill patternType="gray125"/>
            </fill>
          </fills>
          <borders count="1">
            <border/>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="6">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="#{@date_format_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="#{@datetime_format_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="#{@date_format_id}" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="#{@datetime_format_id}" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
          <dxfs count="0"/>
          <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
        </styleSheet>
      XML
    end

    def write_sst
      @writer.add_file "xl/sharedStrings.xml"
      @writer << XML.header
      @writer << %'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{@sst.references}" uniqueCount="#{@sst.size}">'
      @sst.each_key do |string|
        @writer << "<si><t>#{XML.escape_value string}</t></si>"
      end
      @writer << '</sst>'
    end

    def write_workbook_rels
      rid = String.new("rId0")
      @writer.add_file "xl/_rels/workbook.xml.rels"
      @writer << XML.header
      @writer << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      @worksheets.each do |worksheet|
        @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{worksheet.id}.xml"/>'
      end
      @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' unless @sst.empty?
      @writer << '</Relationships>'
    end

    def write_content_types
      @writer.add_file "[Content_Types].xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      XML
      @writer << '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' unless @sst.empty?
      @worksheets.each do |worksheet|
        @writer << %'<Override PartName="/xl/worksheets/sheet#{worksheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      end
      @writer << '</Types>'
    end
  end
end
