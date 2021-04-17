# frozen_string_literal: true

module Xlsxtream
  class Cell
    attr_reader :content, :style

    def initialize(content = nil, style = {})
      @content = content
      @style = style
    end

    def styled?
      !@style.empty?
    end
  end
end
