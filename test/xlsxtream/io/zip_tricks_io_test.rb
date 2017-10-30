require 'test_helper'
require 'xlsxtream/workbook'
require 'xlsxtream/io/hash'
require 'zip'

module Xlsxtream
  class ZipTricksTest < Minitest::Test

    def test_writes_of_multiple_files
      zip_buf = Tempfile.new('ztio-test')

      io = Xlsxtream::IO::ZipTricks.new(zip_buf)
      io.add_file("book1.xml")
      io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />'
      io.add_file("book2.xml")
      io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook>'
      io << '</workbook>'
      io.add_file("empty.txt")
      io.add_file("another.xml")
      io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />'
      io.close

      zip_buf.rewind

      files_contents = {}
      ::Zip::File.open(zip_buf.path) do |zip_file|
        # Handle entries one by one
        zip_file.each do |entry|
          files_contents[entry.name] = entry.get_input_stream.read
        end
      end
      assert_equal '', files_contents['empty.txt']
      assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook></workbook>', files_contents['book2.xml']
      assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />', files_contents['another.xml']
    end
  end
end
