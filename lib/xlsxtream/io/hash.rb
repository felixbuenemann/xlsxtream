module Xlsxtream
  module IO
    class Hash
      def initialize(stream)
        @stream = stream
        @hash = {}
        @path = nil
      end

      def <<(data)
        @stream << data
      end

      def add_file(path)
        close
        @path = path
        @hash[@path] = [@stream.tell]
      end

      def close
        @hash[@path] << @stream.tell if @path
      end

      def fetch(path)
        old = @stream.tell
        from, to = @hash.fetch(path)
        size = to - from
        @stream.seek(from)
        data = @stream.read(size)
        @stream.seek(old)
        data
      end

      def [](path)
        fetch(path)
      rescue KeyError
        nil
      end

      def to_h
        @hash.keys.map {|path| [path, fetch(path)] }.to_h
      end
    end
  end
end

