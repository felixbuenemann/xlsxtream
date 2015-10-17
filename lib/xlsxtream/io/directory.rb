require "pathname"

module Xlsxtream
  module IO
    class Directory
      def initialize(path)
        @path = Pathname(path)
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
        @file.close if @file.respond_to? :close
      end
    end
  end
end
