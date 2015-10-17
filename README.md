# Xlsxtream

Xlsxtream is a streaming writer for XLSX spreadsheets. It supports multiple worksheets and optional string deduplication via a shared string table (SST). Its purpose is to replace CSV for large exports, because using CSV in Excel is very buggy and error prone. It's very efficient and can quickly write millions of rows with low memory usage.

Xlsxtream does not support formatting, charts, comments and a myriad of other [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML) features. If you are looking for a fully featured solution take a look at [axslx](https://github.com/randym/axlsx).

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
# Creates a new workbook and closes it at the end of the block.
XLSXtream::Workbook.open("foo.xlsx") do |xlsx|
  xlsx.write_sheet "Sheet1" do |sheet|
    # Date, Time, DateTime, Fixnum & Float are properly mapped
    sheet << [Date.today, "hello", "world", 42, 3.14159265359]
  end
end

io = StringIO.new('')
xlsx = XLSXtream::Workbook.new(io)
xlsx.write_sheet "Sheet1" do |sheet|
  # Number of columns doesn't have to match
  sheet << %[first row]
  sheet << %[second row with more colums]
end
# Write multiple worksheets with custom names:
xlsx.write_sheet "Foo & Bar" do |sheet|
  sheet.add_row ["Timestamp", "Comment"]
  sheet.add_row [Time.now, "Foo"]
  sheet.add_row [Time.now, "Bar"]
end
# If you have highly repetitive data, you can enable Shared
# String Tables (SST) for the workbook or a single worksheet.
# The SST has to be kept in memory, so don't use it if you
# have a huge amount of rows or a little duplication of content
# accros cells. A single SST is used across the whole workbook.
xlsx.write_sheet("SST", use_shared_strings: true) do |sheet|
  sheet << %(the same old story)
  sheet << %(the old same story)
  sheet << %(old, the same story)
end
# Writes metadata and ZIP archive central directory.
xlsx.close
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/felixbuenemann/xlsxtream.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

