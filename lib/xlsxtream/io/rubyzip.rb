require "zip"

module Xlsxtream
  module IO
    class RubyZip
      def initialize(path_or_io)
        @stream = path_or_io.respond_to? :reopen
        path_or_io.binmode if path_or_io.respond_to? :binmode
        @zos = Zip::OutputStream.new(path_or_io, @stream)
      end

      def <<(data)
        @zos << data
      end

      def add_file(path)
        @zos.put_next_entry path
      end

      def close
        os = @zos.close_buffer
        os.flush if os.respond_to? :flush
        os.close if !@stream and os.respond_to? :close
      end
    end
  end
end
