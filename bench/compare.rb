# frozen_string_literal: true
# Compares the performance of two versions of xlsxtream.
#
#   bundle exec ruby bench/compare.rb [BASE] [TARGET]
#
# BASE and TARGET are git refs. BASE defaults to master, TARGET defaults to
# the working tree (including uncommitted changes). Each ref's lib/ is
# extracted with git archive, and every measurement runs in a fresh Ruby
# process. Runs alternate between the two versions to even out drift.
#
# Environment variables:
#
#   SIZE=large             1,000,000 rows (close to Excel's per-sheet limit),
#                          3 rounds, 1 iteration, streaming only, to check that
#                          large sheets stay fast and memory stays flat
#                          (all scenarios take around 20 minutes)
#   SIZE=xl                10 sheets of 1,000,000 rows, 1 round, streaming
#                          only, mixed and strings_sst_unique by default
#                          (around an hour)
#   ROWS=20000             rows per worksheet
#   SHEETS=1               worksheets per run
#   ROUNDS=5               processes per version and scenario
#   ITERATIONS=3           timed runs per process (the fastest is kept)
#   LAYERS=xml,xlsx        xml: worksheet XML only, xlsx: zipped workbook
#   SOURCES=stream,memory  stream: rows built while writing, memory: rows
#                          built first and written from an array
#   SCENARIOS=mixed,...    defaults to all scenarios in bench/scenarios.rb
#   OPTIONS='{"x":true}'   JSON worksheet options added to every scenario,
#                          e.g. to measure the cost of turning on a new option
#
# Times and allocations exclude building the rows, which is measured
# separately in each process. Allocation counts are deterministic, so any
# change in "allocs/row" is real even when timings are noisy.
require 'fileutils'
require 'json'
require 'open3'
require 'rbconfig'
require 'tmpdir'
require 'zip_kit'
require_relative 'scenarios'

ROOT = File.expand_path('..', __dir__)
BASE = ARGV[0] || 'master'
TARGET = ARGV[1]
SIZES = {
  'default' => { 'ROWS' => '20000', 'SHEETS' => '1', 'ROUNDS' => '5', 'ITERATIONS' => '3',
                 'SOURCES' => 'stream,memory' },
  'large' => { 'ROWS' => '1000000', 'SHEETS' => '1', 'ROUNDS' => '3', 'ITERATIONS' => '1',
               'SOURCES' => 'stream' },
  'xl' => { 'ROWS' => '1000000', 'SHEETS' => '10', 'ROUNDS' => '1', 'ITERATIONS' => '1',
            'SOURCES' => 'stream', 'SCENARIOS' => 'mixed,strings_sst_unique' }
}.freeze
SIZE = SIZES.fetch(ENV.fetch('SIZE', 'default')) { abort "SIZE must be one of #{SIZES.keys.join(', ')}" }
ROWS = Integer(ENV.fetch('ROWS', SIZE['ROWS']))
ROUNDS = Integer(ENV.fetch('ROUNDS', SIZE['ROUNDS']))
SHEETS = Integer(ENV.fetch('SHEETS', SIZE['SHEETS']))
ITERATIONS = Integer(ENV.fetch('ITERATIONS', SIZE['ITERATIONS']))
SOURCES = ENV.fetch('SOURCES', SIZE['SOURCES']).split(',')
LAYERS = ENV.fetch('LAYERS', 'xml,xlsx').split(',')
SCENARIOS = ENV.fetch('SCENARIOS') { SIZE['SCENARIOS'] || Bench::SCENARIOS.keys.join(',') }.split(',')
# Allocations are deterministic, so the first round also counts them for a
# tenth of the rows, to tell a fixed cost from one that grows with rows.
PROBE_ROWS = [ROWS / 10, 1].max
ZIP_KIT_PATHS = Gem.loaded_specs.fetch('zip_kit').full_require_paths
# The summary reports a time or RSS change when it's at least MIN_CHANGE % and
# larger than the noise, since a few rounds can understate the noise. With a
# single round there is no noise to compare against, so changes are reported
# with "noise n/a" for the reader to judge.
MIN_CHANGE = 2.0

def lib_for(ref, dir)
  return File.join(ROOT, 'lib') if ref.nil?

  target = File.join(dir, ref.gsub(/[^\w.-]/, '_'))
  FileUtils.mkdir_p(target)
  archive, status = Open3.capture2('git', '-C', ROOT, 'archive', ref, 'lib', :binmode => true)
  abort "git archive #{ref} failed" unless status.success?
  _, status = Open3.capture2('tar', '-x', '-C', target, :stdin_data => archive, :binmode => true)
  abort "Extracting #{ref} failed" unless status.success?
  File.join(target, 'lib')
end

def measure(lib, scenario, layer, source, probe_rows = nil)
  load_path = ([lib] + ZIP_KIT_PATHS).flat_map { |path| ['-I', path] }
  command = [RbConfig.ruby, *load_path, File.join(__dir__, 'run.rb'),
             scenario, layer, source, ROWS.to_s, SHEETS.to_s, ITERATIONS.to_s, *probe_rows&.to_s]
  # Run without Bundler, which would put the working tree's lib/ (from the
  # gemspec) on the load path ahead of the version being measured.
  env = ENV.to_h.select { |key, _| !key.start_with?('BUNDLE', 'RUBYOPT', 'RUBYLIB') }
  output, status = Open3.capture2(env, *command, :unsetenv_others => true)
  abort "Benchmark run failed: #{command.join(' ')}" unless status.success?

  result = JSON.parse(output)
  unless File.realpath(result['loaded_from']).start_with?(File.realpath(lib))
    abort "Expected xlsxtream from #{lib}, but loaded #{result['loaded_from']}"
  end
  result
end

def median(values)
  sorted = values.sort
  middle = sorted.size / 2
  sorted.size.odd? ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2.0
end

def spread(values)
  (values.max - values.min) / median(values) * 100 if values.size > 1
end

def significant?(change, noise)
  change.abs > [noise || 0, MIN_CHANGE].max
end

def grows?(result)
  result[:difference] != result[:probe][:difference]
end

def mark(flag, blank = ' ')
  flag ? '*' : blank
end

def noise_label(noise)
  noise ? format('%.1f%%', noise) : 'n/a'
end

def range(values, pattern)
  low, high = values.minmax.map { |value| delimit(value, pattern) }
  low == high ? low : "#{low} to #{high}"
end

def delimit(number, pattern = '%d')
  format(pattern, number).gsub(/(\d)(?=(\d{3})+\z)/, '\\1,')
end

def lines(group, results)
  group.size == results.size ? 'every line' : "#{group.size} of #{results.size} lines"
end

def count(number, noun)
  "#{number} #{noun}#{'s' unless number == 1}"
end

# Prints the range of changes for one measure and how many are beyond noise.
def summarize(heading, results, changes, noises, flagged)
  noise = noises.compact.max
  verdict = flagged.empty? ? 'no changes beyond noise' : "#{lines(flagged, results)} beyond noise"
  puts format('  %-8s%s (noise %s, %s)', heading, changes, noise ? format('up to %.1f%%', noise) : 'n/a', verdict)
end

Dir.mktmpdir('xlsxtream-bench') do |dir|
  libs = { :base => lib_for(BASE, dir), :target => lib_for(TARGET, dir) }
  target_name = TARGET || 'working tree'
  started = Process.clock_gettime(Process::CLOCK_MONOTONIC)

  puts "xlsxtream benchmark: #{BASE} (base) vs #{target_name} (target)"
  puts "Ruby #{RUBY_VERSION}, #{SHEETS} x #{ROWS} rows x #{Bench::COLUMNS} columns, " \
       "#{ROUNDS} rounds x #{ITERATIONS} iterations"
  puts "Extra options: #{ENV['OPTIONS']}" if ENV['OPTIONS']
  puts
  puts format('%-18s %-5s %-6s %9s | %9s %9s %7s  %6s | %8s %8s  | %7s %7s',
              '', '', '', 'rows', 'base', 'target', 'change', 'noise',
              'base', 'target', 'base', 'target')
  puts format('%-18s %-5s %-6s %9s | %9s %9s %7s  %6s | %8s %8s  | %7s %7s',
              'scenario', 'layer', 'source', 'ms', 'ms', 'ms', '', '',
              'allocs', 'allocs', 'RSS MB', 'RSS MB')

  results = []
  SCENARIOS.each do |scenario|
    LAYERS.product(SOURCES).each do |layer, source|
      times = { :base => [], :target => [] }
      generate_times = []
      allocations = {}
      rows = nil
      probe = {}
      rss = { :base => [], :target => [] }

      ROUNDS.times do |round|
        order = round.even? ? [:base, :target] : [:target, :base]
        order.each do |version|
          result = measure(libs[version], scenario, layer, source, (PROBE_ROWS if round.zero?))
          times[version] << result['times'].min - result['generate_times'].min
          generate_times << result['generate_times'].min
          allocations[version] = result['allocations'] - result['generate_allocations']
          rows = result['rows']
          rss[version] << result['rss'] / 1024.0
          next unless result['probe_allocations']

          probe[version] = result['probe_allocations']
          probe[:rows] = result['probe_rows']
        end
      end

      base = median(times[:base])
      target = median(times[:target])
      noise = [spread(times[:base]), spread(times[:target])].compact.max
      change = (target - base) / base * 100
      rss_base = median(rss[:base])
      rss_target = median(rss[:target])
      name = format('%-18s %-5s %-6s', scenario, layer, source)
      result = { :name => name, :change => change, :noise => noise,
                 :allocations => allocations, :rows => rows,
                 :difference => allocations[:target] - allocations[:base],
                 :probe => { :difference => probe[:target] - probe[:base], :rows => probe[:rows] },
                 :rss => { :base => rss_base, :target => rss_target },
                 :rss_change => (rss_target - rss_base) / rss_base * 100,
                 :rss_noise => [spread(rss[:base]), spread(rss[:target])].compact.max }
      results << result
      puts format('%s %9.1f | %9.1f %9.1f %+6.1f%%%s %6s | %8.1f %8.1f%s | %7.0f %7.0f%s',
                  name, median(generate_times) * 1000,
                  base * 1000, target * 1000, change, mark(significant?(change, noise)), noise_label(noise),
                  allocations[:base].fdiv(rows), allocations[:target].fdiv(rows), mark(grows?(result)),
                  rss_base, rss_target, mark(significant?(result[:rss_change], result[:rss_noise]), ''))
    end
  end

  puts
  puts 'rows: time spent building the rows while streaming, the same for both versions and excluded'
  puts '      from the rest (0 for the memory source, where rows are built before timing)'
  puts 'base/target ms: median time spent in xlsxtream'
  puts 'change: target vs base (negative is faster)'
  puts 'noise: largest spread (max - min) / median across rounds; ignore changes smaller than this'
  puts '       (n/a with a single round, when only large changes are meaningful)'
  puts 'allocs: average objects allocated by xlsxtream per row (deterministic)'
  puts 'RSS MB: median resident memory at the end of a run across rounds (approximate, varies by a'
  puts '        few MB between identical runs; includes the rows themselves for the memory source)'
  puts '*: worth a closer look: a time or RSS change beyond noise (and at least 2%), or allocations'
  puts '   that grow with rows'

  puts
  puts format('Summary (%s: %s x %s x %s)', count(results.size, 'line'), count(SCENARIOS.size, 'scenario'),
              count(LAYERS.size, 'layer'), count(SOURCES.size, 'source'))
  changes = range(results.map { |result| result[:change] }, '%+.1f%%')
  flagged = results.select { |result| significant?(result[:change], result[:noise]) }
  summarize('Time:', results, changes, results.map { |result| result[:noise] }, flagged)

  allocs = results.reject { |result| result[:difference].zero? }
  growing, fixed = allocs.partition { |result| grows?(result) }
  summary = []
  unless growing.empty?
    summary << format('%s with rows (%s at %s rows, %s at %s rows)',
                      growing.size == results.size ? 'every line grows' : "#{lines(growing, results)} grow",
                      range(growing.map { |result| result[:probe][:difference] }, '%+d'),
                      delimit(growing.first[:probe][:rows]),
                      range(growing.map { |result| result[:difference] }, '%+d'), delimit(growing.first[:rows]))
  end
  unless fixed.empty?
    summary << format('%s objects in %s, at both %s and %s rows (a fixed cost)',
                      range(fixed.map { |result| result[:difference] }, '%+d'), lines(fixed, results),
                      delimit(fixed.first[:probe][:rows]), delimit(fixed.first[:rows]))
  end
  summary << 'identical in every line' if allocs.empty?
  summary.each_with_index { |line, index| puts format('  %-8s%s', index.zero? ? 'Allocs:' : '', line) }

  changes = range(results.map { |result| result[:rss][:target] - result[:rss][:base] }, '%+.1f')
  flagged = results.select { |result| significant?(result[:rss_change], result[:rss_noise]) }
  summarize('Memory:', results, "#{changes} MB", results.map { |result| result[:rss_noise] }, flagged)
  puts format('Finished in %.0fs', Process.clock_gettime(Process::CLOCK_MONOTONIC) - started)
end
