# frozen_string_literal: true
# Runs one benchmark scenario against whichever xlsxtream is first on the
# load path and prints the result as JSON. Normally called by
# bench/compare.rb, but can be run directly:
#
#   bundle exec ruby -Ilib bench/run.rb mixed xml stream 20000 1 5 [2000]
#
# Arguments:
#
#   scenario    name from bench/scenarios.rb
#   layer       "xml" for worksheet XML only, "xlsx" for a zipped workbook
#   source      "stream" builds each row while writing, like a caller reading
#               from a database cursor; "memory" builds every row first and
#               writes from the array, like a caller that loads all records
#   rows        rows per worksheet
#   sheets      number of worksheets (the memory source writes the same rows
#               to each)
#   iterations  timed runs
#   probe rows  optional; also counts allocations when writing this many rows
#               per worksheet, so a caller can tell a fixed cost from one that
#               grows with the number of rows
#
# For the stream source, building the rows without writing them is timed
# separately, so that the gem's share of the time can be reported.
require 'json'
require 'xlsxtream'
require_relative 'scenarios'

# Output sink that discards everything, so no I/O is measured.
class NullIO
  def <<(_data)
    self
  end

  def write(data)
    data.bytesize
  end
end

WARMUP_ROWS = 2_000

name, layer, source, row_count, sheet_count, iterations, probe_rows = ARGV
row_count = Integer(row_count || 20_000)
sheet_count = Integer(sheet_count || 1)
iterations = Integer(iterations || 5)
probe_rows = Integer(probe_rows) if probe_rows

scenario = Bench::SCENARIOS.fetch(name)
build_row = scenario[:row]
options = scenario[:options]
options = options.merge(JSON.parse(ENV['OPTIONS'], :symbolize_names => true)) if ENV['OPTIONS']

each_row =
  case source
  when 'stream'
    # Row numbers carry on across sheets, so every sheet gets new values
    lambda { |count, sheet = 0, &block| count.times { |i| block.call(build_row.call(sheet * count + i)) } }
  when 'memory'
    rows = Array.new(row_count) { |i| build_row.call(i) }
    lambda { |count, _sheet = 0, &block| (count == rows.size ? rows : rows.first(count)).each(&block) }
  else
    abort "Unknown source #{source.inspect}, use stream or memory"
  end

generate = lambda do |count|
  sheet_count.times { |sheet| each_row.call(count, sheet) { |_row| } }
end

write = lambda do |count|
  case layer
  when 'xml'
    worksheet_options = options.dup
    # The workbook normally turns this option into a shared string table
    if worksheet_options.delete(:use_shared_strings)
      worksheet_options[:sst] = Xlsxtream::SharedStringTable.new
    end
    sheet_count.times do |sheet|
      worksheet = Xlsxtream::Worksheet.new(NullIO.new, worksheet_options.dup)
      each_row.call(count, sheet) { |row| worksheet << row }
      worksheet.close
    end
  when 'xlsx'
    Xlsxtream::Workbook.open(NullIO.new) do |workbook|
      sheet_count.times do |sheet|
        workbook.write_worksheet("Sheet#{sheet + 1}", options.dup) do |worksheet|
          each_row.call(count, sheet) { |row| worksheet << row }
        end
      end
    end
  else
    abort "Unknown layer #{layer.inspect}, use xml or xlsx"
  end
end

def measure(callable, count)
  GC.start
  allocated_before = GC.stat(:total_allocated_objects)
  started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
  callable.call(count)
  elapsed = Process.clock_gettime(Process::CLOCK_MONOTONIC) - started
  [elapsed, GC.stat(:total_allocated_objects) - allocated_before]
end

warmup_rows = [row_count, WARMUP_ROWS].min
each_row.call(warmup_rows) { |_row| }
write.call(warmup_rows)

# Measured after the warmup, which includes one-off costs like loading code
if probe_rows
  _, probe_allocations = measure(write, probe_rows)
  probe_allocations -= measure(generate, probe_rows).last if source == 'stream'
end

times = []
generate_times = []
allocations = generate_allocations = nil
iterations.times do
  # Iterating an in-memory array costs next to nothing, so only the stream
  # source needs the cost of building rows subtracted.
  if source == 'stream'
    generate_time, generate_allocations = measure(generate, row_count)
  else
    generate_time, generate_allocations = 0.0, 0
  end
  time, allocations = measure(write, row_count)
  generate_times << generate_time
  times << time
end

# Resident memory at the end of the run, in kB. Approximate, but enough to
# show whether memory grows with the number of rows. For the memory source
# it includes the rows themselves.
rss = Integer(`ps -o rss= -p #{Process.pid}`.strip)
loaded_from = $LOADED_FEATURES.grep(%r{/xlsxtream/worksheet\.rb\z}).first

puts JSON.generate(
  :times => times, :generate_times => generate_times,
  :allocations => allocations, :generate_allocations => generate_allocations,
  :rows => row_count * sheet_count, :rss => rss, :loaded_from => loaded_from,
  :probe_allocations => probe_allocations, :probe_rows => probe_rows && probe_rows * sheet_count
)
