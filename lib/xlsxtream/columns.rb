# encoding: utf-8
require "xlsxtream/xml"

module Xlsxtream
  class Columns

    ENCODING = Encoding.find('UTF-8')

    # Pass an Array of column options Hashes. Symbol Hash keys and associated
    # values are as follows:
    #
    # +width_chars+::  Approximate column with in characters, calculated per
    #                  MSDN docs as if using a default 11 point Calibri font
    #                  for a 96 DPI target. Specify as an integer.
    #
    # +width_pixels+:: Exact with of column in pixels. Specify as a Float.
    #                  Overrides +width_chars+ if that is also provided.
    #
    def initialize(column_options_array)
      @columns = column_options_array
    end

    def to_xml
      xml = '<cols>'

      @columns.each_with_index do |column, index|
        width_chars  = column[ :width_chars  ]
        width_pixels = column[ :width_pixels ]

        if width_chars.nil? && width_pixels.nil?
          xml << %Q{<col min="#{index + 1}" max="#{index + 1}"/>}
        else

          # https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.column.aspx
          #
          # Truncate(
          #   [{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]
          #   /{Maximum Digit Width}*256
          # )/256
          #
          # "Using the Calibri font as an example, the maximum digit width of
          #  11 point font size is 7 pixels (at 96 dpi)"
          #
          width_pixels ||= ((((width_chars * 7.0) + 5) / 7) * 256).truncate() / 256.0

          xml << %Q{<col min="#{index + 1}" max="#{index + 1}" width="#{width_pixels}" customWidth="1"/>}
        end
      end

      xml << '</cols>'
    end

  end
end
