require "zip"

module Xlsxtream
  module IO
    class RubyZip
      def initialize(path_or_io)
        stream = path_or_io.respond_to? :reopen
        path_or_io.binmode if path_or_io.respond_to? :binmode
        @zos = UnbufferedZipOutputStream.new(path_or_io, stream)
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
        os.close if os.respond_to? :close
      end

      # Extend get_compressor to hook our custom deflater.
      class UnbufferedZipOutputStream < ::Zip::OutputStream
        private
        def get_compressor(entry, level)
          case entry.compression_method
          when ::Zip::Entry::DEFLATED then
            StreamingDeflater.new(@output_stream, level, @encrypter)
          else
            super
          end
        end
      end

      # RubyZip's Deflater buffers to a StringIO until finish is called.
      # This StreamingDeflater writes out chunks during compression.
      class StreamingDeflater < ::Zip::Compressor
        def initialize(output_stream, level = Zip.default_compression, encrypter = NullEncrypter.new)
          super()
          @output_stream = output_stream
          @zlib_deflater = ::Zlib::Deflate.new(level, -::Zlib::MAX_WBITS)
          @size          = 0
          @crc           = ::Zlib.crc32
          unless encrypter.is_a? ::Zip::NullEncrypter
            raise ::Zip::Error, 'StreamingDeflater does not support encryption'
          end
        end

        def <<(data)
          val   = data.to_s
          @crc  = Zlib.crc32(val, @crc)
          @size += val.bytesize
          @output_stream << @zlib_deflater.deflate(data)
        end

        def finish
          @output_stream << @zlib_deflater.finish until @zlib_deflater.finished?
        end

        attr_reader :size, :crc
      end
    end
  end
end
