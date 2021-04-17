# frozen_string_literal: true

module Xlsxtream
  module Styles
    class Border
      def to_xml
        '<border/>'
      end

      def ==(other)
        self.class == other.class && state == other.state
      end
      alias eql? ==

      def hash
        state.hash
      end

      def state
        []
      end
    end
  end
end
