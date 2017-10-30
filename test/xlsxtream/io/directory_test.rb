require 'test_helper'
require 'xlsxtream/io/directory'
require 'pathname'

module Xlsxtream
  module IO
    class DirectoryTest < Minitest::Test

      def test_writes_of_multiple_files
        Dir.mktmpdir do |dir|
          io = Xlsxtream::IO::Directory.new(dir)
          io.add_file("book1.xml")
          io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />'
          io.add_file("book2.xml")
          io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook>'
          io << '</workbook>'
          io.add_file("empty.txt")
          io.add_file("another.xml")
          io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />'
          io.close

          dir = Pathname(dir)
          assert_equal '', dir.join('empty.txt').read
          assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />', dir.join('book1.xml').read
          assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook></workbook>', dir.join('book2.xml').read
          assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />', dir.join('another.xml').read
        end

      end
    end
  end
end
