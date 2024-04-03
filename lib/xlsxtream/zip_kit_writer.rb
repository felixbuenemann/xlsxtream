# frozen_string_literal: true
require "zip_kit"

module Xlsxtream
  class ZipKitWriter
    BUFFER_SIZE = 64 * 1024

    def self.with_output_to(output)
      if output.is_a?(self)
        output
      elsif output.is_a?(ZipKit::Streamer)
        # If this is a Streamer which has already been initialized, it is likely that the streamer
        # was initialized with a Streamer.open block - it will close itself. This allows xslxstream
        # to be used with zip_kit_stream and other cases where the Streamer is managed externally
        new(output, close: [])
      elsif output.is_a?(String)
        file = File.open(output, 'wb')
        streamer = ZipKit::Streamer.new(file)
        # First the Streamer needs to be closed (to write out the central directory), then the file
        new(streamer, close: [streamer, file])
      elsif output.respond_to?(:<<) || output.respond_to?(:write)
        streamer = ZipKit::Streamer.new(output)
        new(streamer, close: [streamer])
      else
        error = <<~MSG
          An `output` object must be one of:

          * A String containing a path to a file ("workbook.xslx")
          * A ZipKit::Streamer
          * An IO-like object responding to #<< or #write

          but it was a #{output.class}
        MSG
        raise ArgumentError, error
      end
    end

    def initialize(streamer, close: [])
      @streamer = streamer
      @currently_writing_file_inside_zip = nil
      @buffer = String.new
      @close = close
    end

    def <<(data)
      @buffer << data
      flush_buffer if @buffer.size >= BUFFER_SIZE
      self
    end

    def add_file(path)
      flush_file
      @currently_writing_file_inside_zip = @streamer.write_deflated_file(path)
    end

    def close
      flush_file
      @close.each(&:close)
    end

    private

    def flush_buffer
      @currently_writing_file_inside_zip << @buffer
      @buffer.clear
    end

    def flush_file
      return unless @currently_writing_file_inside_zip
      flush_buffer if @buffer.size > 0
      @currently_writing_file_inside_zip.close
    end
  end
end
