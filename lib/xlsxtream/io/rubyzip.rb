# frozen_string_literal: true
require "zip"
require "xlsxtream/errors"

module Xlsxtream
  module IO
    class RubyZip
      def initialize(io)
        unless io.respond_to? :pos and io.respond_to? :pos=
          raise Error, 'IO is not seekable'
        end
        io.binmode if io.respond_to? :binmode
        stream = true
        @zos = Zip::OutputStream.new(io, stream)
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
      end
    end
  end
end
