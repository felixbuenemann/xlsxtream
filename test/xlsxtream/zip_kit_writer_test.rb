# frozen_string_literal: true
require 'test_helper'
require 'zip'

module Xlsxtream
  class ZipKitWriterTest < Minitest::Test
    def test_writes_of_multiple_files
      zip_buf = Tempfile.new('ztio-test')

      io = Xlsxtream::ZipKitWriter.with_output_to(zip_buf)
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

    def test_with_output_to_wraps_another_writer
      another_writer = Xlsxtream::ZipKitWriter.new(ZipKit::Streamer.new(StringIO.new))
      assert_equal another_writer, Xlsxtream::ZipKitWriter.with_output_to(another_writer)
    end

    def test_with_output_to_wraps_a_zip_kit_streamer_and_does_not_close_it_on_close
      streamer_that_raises = Class.new(ZipKit::Streamer) do
        def close
          raise "Should not happen"
        end
      end.new(StringIO.new)

      writer = Xlsxtream::ZipKitWriter.new(streamer_that_raises)
      writer.close
    end

    def test_with_output_to_creates_a_file_with_a_given_path
      Dir.mktmpdir do |dir_path|
        tf_name = "output.xlsx"
        writer = Xlsxtream::ZipKitWriter.with_output_to(File.join(dir_path, tf_name))
        assert File.exist?(File.join(dir_path, tf_name))
        writer.close

        assert File.exist?(File.join(dir_path, tf_name))
      end
    end

    def test_with_output_to_converts_pathname_into_path
      tf = Tempfile.new
      assert tf.size == 0
      pathname = Pathname.new(tf.path)
      writer = Xlsxtream::ZipKitWriter.with_output_to(pathname)
      writer.close
      assert tf.size > 0
    end

    def test_with_output_to_writes_into_io
      io = StringIO.new
      writer = Xlsxtream::ZipKitWriter.with_output_to(io)
      writer.close
      assert io.size > 0
    end

    def test_with_output_to_raises_on_unwritable_arg
      assert_raises ArgumentError do
        Xlsxtream::ZipKitWriter.with_output_to(:sym)
      end
    end
  end
end
