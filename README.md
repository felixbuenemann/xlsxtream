# Xlsxtream

[![Gem Version](https://badge.fury.io/rb/xlsxtream.svg)](https://rubygems.org/gems/xlsxtream)

Xlsxtream is a streaming writer for XLSX spreadsheets. It supports multiple worksheets and optional string
deduplication via a shared string table (SST). Its purpose is to replace CSV for large exports, because using
CSV in Excel is very buggy and error prone. It's very efficient and can quickly write millions of rows with
low memory usage.

Xlsxtream does not support formatting, charts, comments and a myriad of
other [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML) features. If you are looking for a
fully featured solution take a look at [caxslx](https://github.com/caxlsx/caxlsx).

Xlsxtream supports writing to files or IO-like objects, data is flushed as the ZIP compressor sees fit.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'xlsxtream'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsxtream

## Usage

```ruby
# Creates a new workbook file, write and close it at the end of the block
Xlsxtream::Workbook.open('my_data.xlsx') do |xlsx|
  xlsx.write_worksheet 'Sheet1' do |sheet|
    # Boolean, Date, Time, DateTime and Numeric are properly mapped
    sheet << [true, Date.today, 'hello', 'world', 42, 3.14159265359, 42**13]
  end
end

io = StringIO.new
xlsx = Xlsxtream::Workbook.new(io)

# Number of columns doesn't have to match
xlsx.write_worksheet 'Sheet1' do |sheet|
  sheet << ['first', 'row']
  sheet << ['second', 'row', 'with', 'more colums']
end

# Write multiple worksheets with custom names
xlsx.write_worksheet 'AppendixSheet' do |sheet|
  sheet.add_row ['Timestamp', 'Comment']
  sheet.add_row [Time.now, 'Good times']
  sheet.add_row [Time.now, 'Time-machine']
end

# If you have highly repetitive data, you can enable Shared String Tables (SST)
# for the workbook or a single worksheet. The SST has to be kept in memory,
# so do not use it if you have a huge amount of rows or a little duplication
# of content across cells. A single SST is used for the whole workbook.
xlsx.write_worksheet(name: 'SheetWithSST', use_shared_strings: true) do |sheet|
  sheet << ['the', 'same', 'old', 'story']
  sheet << ['the', 'old', 'same', 'story']
  sheet << ['old', 'the', 'same', 'story']
end

# Strings in numeric or date/time format can be auto-detected and formatted
# appropriately. This is a convenient way to avoid an Excel-warning about
# "Number stored as text". Dates and times must be in the ISO-8601 format and
# numeric values must contain only numbers and an optional decimal separator.
# The strings true and false are detected as boolean values.
xlsx.write_worksheet(name: 'SheetWithAutoFormat', auto_format: true) do |sheet|
  # these two rows will be identical in the xlsx-output
  sheet << [true, 11.85, DateTime.parse('2050-01-01T12:00'), Date.parse('1984-01-01')]
  sheet << ['true', '11.85', '2050-01-01T12:00', '1984-01-01']
end

# You can also create worksheet without a block, using the `add_worksheet` method.
# It can be only used sequentially, so remember to manually close the worksheet
# when you are done (before opening a new one).
worksheet = xls.add_worksheet(name: 'SheetWithoutBlock')
worksheet << ['some', 'data']
worksheet.close

# Writes metadata and ZIP archive central directory
xlsx.close
# Close IO object
io.close

# Changing the default font from Calibri, 12pt, Swiss
Xlsxtream::Workbook.new(io, font: {
  name: 'Times New Roman',
  size: 10, # size in pt
  family: 'Roman' # Swiss, Modern, Script, Decorative
})

# Specifying column widths in pixels or characters; 3 column example;
# "pixel" widths appear to be *relative* to an assumed 11pt Calibri
# font, so if selecting a different font or size (see above), do not
# adjust widths to match. Calculate pixel widths for 11pt Calibri.
Xlsxtream::Workbook.new(io, columns: [
  { width_pixels: 33 },
  { width_chars: 7 },
  { width_chars: 24 }
])
# The :columns option can also be given to write_worksheet, so it's
# possible to have multiple worksheets with different column widths.


# Output from Rails (will stream without buffering)
class ReportsController < ApplicationController
  include ZipKit::RailsStreaming
  EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

  def download
    zip_kit_stream(filename: "report.xlsx", type: EXCEL_CONTENT_TYPE) do |zip_kit_streamer|
      Xlsxtream::Workbook.open(zip_kit_streamer) do |xlsx|
        xlsx.write_worksheet 'Sheet1' do |sheet|
          # Boolean, Date, Time, DateTime and Numeric are properly mapped
          sheet << [true, Date.today, 'hello', 'world', 42, 3.14159265359, 42**13]
        end
      end
    end
  end
end
```

## Compatibility

The current version of Xlsxtream requires at least Ruby 2.6

If you are using an older Ruby version you can use the following in your Gemfile:

```ruby
gem 'xlsxtream', '< 3'
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/felixbuenemann/xlsxtream.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
