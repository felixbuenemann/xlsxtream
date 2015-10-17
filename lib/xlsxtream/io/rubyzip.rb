require "zip"

module Xlsxtream
  module IO
    class RubyZip
      def initialize(path_or_io)
        stream = path_or_io.respond_to? :reopen
        @zos = ::Zip::OutputStream.new(path_or_io, stream)
      end

      def <<(data)
        @zos << data
      end

      def add_file(path)
        @zos.put_next_entry path
      end

      def close
        @zos.close
      end
    end
  end
end
