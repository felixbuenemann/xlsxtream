require "zip_tricks"

module Xlsxtream
  module IO
    class ZipTricksFibers
      def initialize(body)
        @streamer = ::ZipTricks::Streamer.new(body)
        @wf = nil
      end

      def <<(data)
        @wf.resume(data)
        self
      end

      def add_file(path)
        @wf.resume(:__close__) if @wf

        @wf = Fiber.new do | bytes_or_close_sym |
          @streamer.write_deflated_file(path) do |write_sink|
            loop do
              break if bytes_or_close_sym == :__close__
              write_sink << bytes_or_close_sym
              bytes_or_close_sym = Fiber.yield
            end
          end
        end
      end

      def close
        @wf.resume(:__close__) if @wf
        @streamer.close
      end
    end
  end
end
