# frozen_string_literal: true
require "pathname"

module Xlsxtream
  module IO
    class Directory
      def initialize(path)
        @path = Pathname(path)
        @file = nil
      end

      def <<(data)
        @file << data
      end

      def add_file(path)
        close
        file_path = @path + path
        file_path.parent.mkpath
        @file = file_path.open("wb")
      end

      def close
        @file.close if @file
      end
    end
  end
end
