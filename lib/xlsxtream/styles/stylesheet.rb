# frozen_string_literal: true

require 'set'

require "xlsxtream/styles/border"
require "xlsxtream/styles/fill"
require "xlsxtream/styles/font"
require "xlsxtream/styles/xf"

module Xlsxtream
  module Styles
    class Stylesheet
      # TODO: create a class for NumFormat to replace the constants below
      DEFAULT_NUM_FORMAT_ID = 0
      DATE_NUM_FORMAT_ID = 164
      TIME_NUM_FORMAT_ID = 165
      SUPPORTED_GLOBAL_OPTIONS = Set[:font, :fill]

      # Parameter `global_options` - represents a hash of default options to be
      # applied to the whole book, for example:
      #   `global_options = { font: { size: 10 }, fill: { color: 'FFFF00' } }`
      def initialize(global_options = {})
        @global_options = extract_styles_from global_options

        @borders = Hash.new { |h, k| h[k] = h.size }
        @fills   = Hash.new { |h, k| h[k] = h.size }
        @fonts   = Hash.new { |h, k| h[k] = h.size }
        @xfs     = Hash.new { |h, k| h[k] = h.size }

        populate_with_basic_xfs!
      end

      def style_id(cell)
        xf_args = {}

        xf_args[:numFmtId] =  case cell.content
                              when Date           then DATE_NUM_FORMAT_ID
                              when Time, DateTime then TIME_NUM_FORMAT_ID
                              else DEFAULT_NUM_FORMAT_ID
                              end

        xf_args[:borderId] = @borders[Border.new]
        xf_args[:fontId] = @fonts[Font.new(cell.style[:font] || {})]
        xf_args[:fillId] = @fills[Fill.new(cell.style[:fill] || {})]
        xf_args[:applyFill] = 1 unless xf_args[:fillId] == @fills[default_fill]

        @xfs[Xf.new(xf_args)]
      end

      def to_xml
        XML.header + XML.strip(<<-XML)
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="2">
              <numFmt numFmtId="#{DATE_NUM_FORMAT_ID}" formatCode="yyyy\\-mm\\-dd"/>
              <numFmt numFmtId="#{TIME_NUM_FORMAT_ID}" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>
            </numFmts>
            <fonts count="#{@fonts.size}">#{@fonts.keys.map(&:to_xml).join}</fonts>
            <fills count="#{@fills.size}">#{@fills.keys.map(&:to_xml).join}</fills>
            <borders count="#{@borders.size}">#{@borders.keys.map(&:to_xml).join}</borders>
            <cellStyleXfs count="1">#{@xfs.keys.first.to_xml}</cellStyleXfs>
            <cellXfs count="#{@xfs.size}">#{@xfs.keys.map(&:to_xml).join}</cellXfs>
            <cellStyles count="1">
              <cellStyle name="Normal" xfId="0" builtinId="0"/>
            </cellStyles>
            <dxfs count="0"/>
            <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
          </styleSheet>
        XML
      end

      def default_style_id
        @default_style_id ||= @xfs[Xf.new(
          numFmtId: DEFAULT_NUM_FORMAT_ID,
          fontId: @fonts[default_font],
          fillId: @fills[default_fill],
          borderId: @borders[default_border]
        )]
      end

      def datetime_style_id
        @datetime_style_id ||= @xfs[Xf.new(
          numFmtId: TIME_NUM_FORMAT_ID,
          fontId: @fonts[default_font],
          fillId: @fills[default_fill],
          borderId: @borders[default_border],
          applyNumberFormat: 1
        )]
      end

      def date_style_id
        @date_style_id ||= @xfs[Xf.new(
          numFmtId: DATE_NUM_FORMAT_ID,
          fontId: @fonts[default_font],
          fillId: @fills[default_fill],
          borderId: @borders[default_border],
          applyNumberFormat: 1
        )]
      end

      private

      def default_border
        # no options for `Border` are supported currently
        @default_border ||= Border.new
      end

      def default_fill
        @default_fill ||= begin
          @global_options.key?(:fill) ? Fill.new(@global_options[:fill]) : Fill.new
        end
      end

      def default_font
        @default_font ||= begin
          @global_options.key?(:font) ? Font.new(@global_options[:font]) : Font.new
        end
      end

      def populate_with_basic_xfs!
        # add default entities to the collections
        @borders[default_border]
        @fills[default_fill]
        @fonts[default_font]

        # populate xfs
        default_style_id
        date_style_id
        datetime_style_id
      end

      def extract_styles_from(options)
        SUPPORTED_GLOBAL_OPTIONS.reduce({}) do |result, option_key|
          result[option_key] = options[option_key] if options[option_key].is_a?(Hash)
          result
        end
      end
    end
  end
end
