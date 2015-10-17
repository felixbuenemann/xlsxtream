module Xlsxtream
  module IO
    class Stream
      def initialize(stream)
        @stream = stream
      end

      def <<(data)
        @stream << data
      end

      def add_file(path)
        close
        @path = path
        @stream << "#@path\n"
      end

      def close
        @stream << "\n" if @path
      end
    end
  end
end

