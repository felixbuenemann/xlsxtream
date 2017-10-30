require 'test_helper'
require 'xlsxtream/io/stream'

module Xlsxtream
  module IO
    class StreamTest < Minitest::Test

      def test_writes_of_multiple_files
        buffer = StringIO.new

        io = Xlsxtream::IO::Stream.new(buffer)
        io.add_file("book1.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />'
        io.add_file("book2.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook>'
        io << '</workbook>'
        io.add_file("empty.txt")
        io.add_file("another.xml")
        io << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />'
        io.close

        file_contents = buffer.string
        assert_equal <<-EOF, file_contents
book1.xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook />
book2.xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook></workbook>
empty.txt

another.xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><another />
EOF
      end
    end
  end
end
