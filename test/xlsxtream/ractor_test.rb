# frozen_string_literal: true
require 'test_helper'
require 'stringio'
require 'zip'

module Xlsxtream
  class RactorTest < Minitest::Test
    def self.write_workbook
      io = StringIO.new(String.new)
      Workbook.open(io, font: { name: 'Arial', family: 'Roman' }, columns: [{ width_chars: 12 }, {}]) do |wb|
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
      io.string
    end

    def test_write_workbook_in_ractor
      skip 'Ractor is not available' unless defined?(Ractor)

      # Writing in the main Ractor first also loads the classes that zip_kit
      # autoloads, which Ruby before 3.4 cannot do from another Ractor.
      expected = unzip(RactorTest.write_workbook)

      ractor = new_ractor { RactorTest.write_workbook }
      actual = unzip(ractor.respond_to?(:value) ? ractor.value : ractor.take)

      assert_equal expected, actual
      assert_includes actual['xl/sharedStrings.xml'], '<si><t>The _x0007_ rings</t></si>'
      assert_includes actual['xl/sharedStrings.xml'], '<si><t>_x005f_x0041_</t></si>'
      assert_includes actual['xl/worksheets/sheet2.xml'], '<is><t>The _x0007_ rings</t></is>'
    end

    private

    def new_ractor(&block)
      experimental = Warning[:experimental]
      Warning[:experimental] = false
      Ractor.new(&block)
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
