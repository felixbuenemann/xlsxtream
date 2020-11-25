# frozen_string_literal: true

module Xlsxtream
  class HeaderRow < Row
    def initialize(row, rownum, options = {})
      super

      @normal_style = ' s="3"'
      @date_style = ' s="4"'
      @time_style = ' s="5"'
    end
  end
end
