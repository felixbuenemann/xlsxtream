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
        # was initialized with a Streamer.open block - it will close itself
        new(output, close: [])
      elsif output.is_a?(String) || !output.respond_to?(:<<)
        @file = File.open(output, 'wb')
        streamer = ZipKit::Streamer.new(@file)
        new(streamer, close: [streamer, @file])
      elsif output.respond_to?(:<<) || output.respond_to?(:write)
        streamer = ZipKit::Streamer.new(output)
        new(streamer, close: [streamer])
      else
        error = <<~MSG
          An `output` object must be one of:

          * A String containing a path to a file ("workbook.xslx")
          * A ZipKit::Streamer
          * An IO-like object responding to #<< or #write

          but it was an #{output.class}
        MSG
        raise ArgumentError, error
      end
    end

    def initialize(streamer, close: [])
      @streamer = streamer
      @wf = nil
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
      @wf = @streamer.write_deflated_file(path)
    end

    def close
      flush_file
      @close.each(&:close)
    end

    private

    def flush_buffer
      @wf << @buffer
      @buffer.clear
    end

    def flush_file
      return unless @wf
      flush_buffer if @buffer.size > 0
      @wf.close
    end
  end
end
