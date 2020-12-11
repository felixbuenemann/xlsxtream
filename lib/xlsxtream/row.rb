# frozen_string_literal: true

require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    TRUE_STRING = 'true'
    FALSE_STRING = 'false'
    GENERIC_STRING_TYPE = 'str' # can be used for cells containing formulas
    SHARED_STRING_TYPE = 's' # used only for shared strings
    INLINE_STRING_TYPE = 'inlineStr' # without formulas, can be of rich text format
    NUMERIC_TYPE = 'n'
    BOOLEAN_TYPE = 'b'
    V_WRAPPER   = ->(text) { "<v>#{text}</v>" }
    IST_WRAPPER = ->(text) { "<is><t>#{text}</t></is>" }

    def initialize(row, rownum, workbook, options = {})
      @row = row
      @rownum = rownum
      @stylesheet = workbook.stylesheet
      @sst = options[:sst]
      @auto_format = options[:auto_format]
    end

    def to_xml
      column = String.new('A')
      xml = String.new(%Q{<row r="#{@rownum}">})

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!

        value = auto_format(value) if @auto_format && value.is_a?(String)
        content, type = prepare_cell_content_and_resolve_type(value)

        # no xml output for empty non-styled strings
        next if content.nil? && !(value.respond_to?(:styled?) && value.styled?)

        style = resolve_cell_style(value)

        # The code below renders a single XML cell.
        #
        # As Xlsxtream library is optimized for performance and memory, it was decided to keep
        # the cell rendering logic here and not in the `Cell` class despite OOP standards encourage
        # otherwise. This is to avoid unnecessary memory allocations in case of low share of
        # non-styled cell content (with no need to use `Cell` wrapper).
        xml << %{<c r="#{cid}"}
        xml << %{ t="#{type}"} if type
        xml << %{ s="#{style}"} if style
        xml << '>'
        xml << (type == INLINE_STRING_TYPE ? IST_WRAPPER : V_WRAPPER)[content]
        xml << '</c>'
      end

      xml << '</row>'
    end

    private

    def prepare_cell_content_and_resolve_type(value)
      case value
      when Numeric
        [value, NUMERIC_TYPE]
      when TrueClass, FalseClass
        [(value ? 1 : 0), BOOLEAN_TYPE]
      when Time
        [time_to_oa_date(value), nil]
      when DateTime
        [datetime_to_oa_date(value), nil]
      when Date
        [date_to_oa_date(value), nil]
      when Cell
        prepare_cell_content_and_resolve_type(value.content)
      else
        value = value.to_s
        return [nil, nil] if value.empty?

        value = value.encode(ENCODING) if value.encoding != ENCODING

        if @sst
          [@sst[value], SHARED_STRING_TYPE]
        else
          [XML.escape_value(value), INLINE_STRING_TYPE]
        end
      end
    end

    def resolve_cell_style(value)
      case value
      when Time, DateTime then @stylesheet.datetime_style_id
      when Date           then @stylesheet.date_style_id
      when Cell           then @stylesheet.style_id(value)
      end
    end

    # Detects and casts numbers, date, time in text
    def auto_format(value)
      case value
      when TRUE_STRING
        true
      when FALSE_STRING
        false
      when NUMBER_PATTERN
        value.include?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        Date.parse(value) rescue value
      when TIME_PATTERN
        DateTime.parse(value) rescue value
      else
        value
      end
    end

    # Converts Time instance to OLE Automation Date
    def time_to_oa_date(time)
      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      (time.to_f + time.utc_offset) / 86400 + 25569
    end

    # Converts DateTime instance to OLE Automation Date
    if RUBY_ENGINE == 'ruby'
      def datetime_to_oa_date(date)
        _, jd, df, sf, of = date.marshal_dump
        jd - 2415019 + (df + of + sf / 1e9) / 86400
      end
    else
      def datetime_to_oa_date(date)
        date.jd - 2415019 + (date.hour * 3600 + date.sec + date.sec_fraction.to_f) / 86400
      end
    end

    # Converts Date instance to OLE Automation Date
    def date_to_oa_date(date)
      (date.jd - 2415019).to_f
    end
  end
end
