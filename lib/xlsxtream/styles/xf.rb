# frozen_string_literal: true

module Xlsxtream
  module Styles
    class Xf
      REQUIRED_FIELDS = %i[
        numFmtId
        fontId
        fillId
        borderId
        xfId
      ].freeze

      OPTIONAL_FIELDS = %i[
        applyFill
        applyNumberFormat
      ].freeze

      SUPPORTED_FIELDS = REQUIRED_FIELDS + OPTIONAL_FIELDS

      def initialize(attrs)
        @attrs = filter attrs
      end

      def to_xml
        tag_attrs = @attrs.map { |k, v| %{#{k}="#{v}"} }.join(' ')

        %{<xf xfId="0" #{tag_attrs}/>}
      end

      def ==(other)
        self.class == other.class && state == other.state
      end
      alias eql? ==

      def hash
        state.hash
      end

      def state
        @attrs
      end

      private

      def filter(attrs)
        attrs.select { |key| SUPPORTED_FIELDS.include?(key.to_sym) }
      end
    end
  end
end
