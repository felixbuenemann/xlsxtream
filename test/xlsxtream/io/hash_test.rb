require 'test_helper'
require 'stringio'
require 'xlsxtream/io/hash'

module Xlsxtream
  module IO
    class HashTest < Minitest::Test

      def test_writes_of_multiple_files
        buffer = StringIO.new

        io = Xlsxtream::IO::Hash.new(buffer)
        io.add_file("book1.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />'
        io.add_file("book2.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook>'
        io << '</workbook>'
        io.add_file("empty.txt")
        io.add_file("another.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />'
        io.close

        file_contents = io.to_h
        assert_equal '', file_contents['empty.txt']
        assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />', file_contents['book1.xml']
        assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook></workbook>', file_contents['book2.xml']
        assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />', file_contents['another.xml']
      end
    end
  end
end
