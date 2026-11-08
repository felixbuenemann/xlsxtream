# frozen_string_literal: true
require 'test_helper'
require 'pathname'
require 'stringio'
require 'tmpdir'
require 'zip'

module Xlsxtream
  class RactorTest < Minitest::Test
    def self.write_workbook(output)
      Workbook.open(output, font: { name: 'Arial', family: 'Roman' }, columns: [{ width_chars: 12 }, {}]) do |wb|
        wb.write_worksheet('Shared & <strings>', use_shared_strings: true, auto_format: true, has_header_row: true) do |ws|
          ws << ['Name', 'Value']
          ws << ["The \x07 rings", '2024-12-28']
          ws << ['_x0041_', '2024-12-28T10:30:00Z']
          ws << ['true', '1.5']
        end
        wb.write_worksheet('Inline') do |ws|
          ws << ["The \x07 rings", 42, true, Date.new(2024, 12, 28), DateTime.new(2024, 12, 28, 10, 30), Time.utc(2024, 12, 28, 10, 30)]
        end
      end
    end

    def setup
      skip 'Ractor is not available' unless defined?(Ractor)

      # Writing in the main Ractor first also loads the classes that zip_kit
      # autoloads, which Ruby before 3.4 cannot do from another Ractor.
      io = StringIO.new(String.new)
      RactorTest.write_workbook(io)
      @expected = unzip(io.string)
    end

    def test_write_workbook_to_io_in_ractor
      xlsx = ractor_value do
        io = StringIO.new(String.new)
        RactorTest.write_workbook(io)
        io.string
      end

      assert_workbook unzip(xlsx)
    end

    def test_write_workbook_to_path_in_ractor
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'workbook.xlsx')
        ractor_value(path) { |file_name| RactorTest.write_workbook(Pathname.new(file_name)) }

        assert_workbook unzip(File.binread(path))
      end
    end

    private

    def assert_workbook(actual)
      assert_equal @expected, actual
      assert_includes actual['xl/sharedStrings.xml'], '<si><t>The _x0007_ rings</t></si>'
      assert_includes actual['xl/sharedStrings.xml'], '<si><t>_x005f_x0041_</t></si>'
      assert_includes actual['xl/worksheets/sheet2.xml'], '<is><t>The _x0007_ rings</t></is>'
    end

    def ractor_value(*args, &block)
      experimental = Warning[:experimental]
      Warning[:experimental] = false
      ractor = Ractor.new(*args, &block)
      ractor.respond_to?(:value) ? ractor.value : ractor.take
    ensure
      Warning[:experimental] = experimental
    end

    def unzip(xlsx)
      file_contents = {}
      Zip::File.open_buffer(StringIO.new(xlsx)).each do |entry|
        file_contents[entry.name] = entry.get_input_stream.read
      end
      file_contents
    end
  end
end
