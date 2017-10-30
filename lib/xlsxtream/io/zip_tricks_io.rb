require "zip_tricks"

module Xlsxtream
  module IO
    class ZipTricksIO
      def initialize(body)
        @streamer = ZipTricks::Streamer.new(body)
      end

      def <<(data)
        @wf << data
        self
      end

      def add_file(path)
        @wf.close if @wf
        @wf = @streamer.write_deflated_file(path)
      end

      def close
        @wf.close if @wf
        @streamer.close
      end
    end
  end
end
