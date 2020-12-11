# frozen_string_literal: true

module Xlsxtream
  module Styles
    class Font
      FAMILY_IDS = {
        ''           => 0,
        'roman'      => 1,
        'swiss'      => 2,
        'modern'     => 3,
        'script'     => 4,
        'decorative' => 5
      }.freeze

      DEFAULT_UNDERLINE = 'single'
      SUPPORTED_UNDERLINES = [
        DEFAULT_UNDERLINE,
        'singleAccounting',
        'double',
        'doubleAccounting'
      ]

      def initialize(bold: nil,
                     italic: nil,
                     strike: nil,
                     underline: nil,
                     color: nil,
                     size: 12,
                     name: 'Calibri',
                     family: 'Swiss')
        @bold = bold
        @italic = italic
        @strike = strike
        @underline = resolve_underline(underline)
        @color = color
        @size = size
        @name = name
        @family_id = resolve_family_id(family)
      end

      def to_xml
        "<font>#{tags.join}</font>"
      end

      def ==(other)
        self.class == other.class && state == other.state
      end
      alias eql? ==

      def hash
        state.hash
      end

      def state
        [@bold, @italic, @strike, @underline, @color, @size, @name, @family_id]
      end

      private

      def tags
        [
          %{<sz val="#{@size}"/>},
          %{<name val="#{@name}"/>},
          %{<family val="#{@family_id}"/>}
        ].tap do |arr|
          arr << %{<b val="true"/>} if @bold
          arr << %{<i val="true"/>} if @italic
          arr << %{<u val="#{@underline}"/>} if @underline
          arr << %{<strike val="true"/>} if @strike
          arr << %{<color rgb="#{@color}"/>} if @color
        end
      end

      def resolve_family_id(value)
        FAMILY_IDS[value.to_s.downcase] or fail Error,
          "Invalid font family #{value}, must be one of "\
          + FAMILY_IDS.keys.map(&:inspect).join(', ')
      end

      def resolve_underline(value)
        return value if SUPPORTED_UNDERLINES.include?(value)
        return DEFAULT_UNDERLINE if value == true
      end
    end
  end
end
