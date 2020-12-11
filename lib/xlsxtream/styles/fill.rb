# frozen_string_literal: true

module Xlsxtream
  module Styles
    class Fill
      NONE = 'none'
      SOLID = 'solid'

      # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-2.8.1
      MS_SUPPORTED = Set[
        NONE,
        SOLID,
        'darkDown',
        'darkGray',
        'darkGrid',
        'darkHorizontal',
        'darkTrellis',
        'darkUp',
        'darkVertical',
        'gray0625',
        'gray125',
        'lightDown',
        'lightGray',
        'lightGrid',
        'lightHorizontal',
        'lightTrellis',
        'lightUp',
        'lightVertical',
        'mediumGray'
      ]

      def initialize(pattern: nil, color: nil)
        @pattern = pattern || (color ? SOLID : NONE)
        @color = color
      end

      def to_xml
        "<fill>#{pattern_tag}</fill>"
      end

      def ==(other)
        self.class == other.class && state == other.state
      end
      alias eql? ==

      def hash
        state.hash
      end

      def state
        [@pattern, @color]
      end

      private

      def color_tag
        return unless @color
        %{<fgColor rgb="#{@color}"/>}
      end

      def pattern_tag
        return unless color_tag
        %{<patternFill patternType="#{@pattern}">#{color_tag}</patternFill>}
      end
    end
  end
end
