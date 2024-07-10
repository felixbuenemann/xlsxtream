# frozen_string_literal: true
require 'test_helper'
require 'xlsxtream/io/zip_kit'
require 'zip'

module Xlsxtream
  class ZipKitTest < Minitest::Test

    def test_writes_of_multiple_files
      zip_buf = Tempfile.new('ztio-test')

      io = Xlsxtream::IO::ZipKit.new(zip_buf)
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

      file_contents = {}
      Zip::File.open(zip_buf) do |zip_file|
        zip_file.each do |entry|
          file_contents[entry.name] = entry.get_input_stream.read
        end
      end
      assert_equal '', file_contents['empty.txt']
      assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook></workbook>', file_contents['book2.xml']
      assert_equal '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />', file_contents['another.xml']
    end
  end
end
