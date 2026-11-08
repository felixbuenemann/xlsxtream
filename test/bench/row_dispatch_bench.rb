# frozen_string_literal: true
# Times Row#to_xml for rows of different shapes and prints {"workload" => seconds}
# as JSON. Meant to be run through test/bench/compare.rb, but works on its own:
#
#   ruby -Ilib test/bench/row_dispatch_bench.rb
require 'json'
require 'date'
if ENV['ACTIVE_SUPPORT']
  require 'active_support'
  require 'active_support/core_ext'
end
require 'xlsxtream/row'

ROWS = Integer(ENV.fetch('ROWS', '50000'))
ITERATIONS = 3

rng = Random.new(42)
words = Array.new(12) { Array.new(8) { ('a'..'z').to_a.sample(:random => rng) }.join }

WORKLOADS = {
  # The shape from PR #55, typical of ActiveRecord exports with sparse columns
  'strings_and_nils' => ['foo', nil, 23] * 4,
  'strings' => words,
  'nils' => [nil] * 12,
  'numbers' => Array.new(12) { |i| i.even? ? rng.rand(9999) : rng.rand(9999.0) },
  'dates_times' => Array.new(12) { |i| i.even? ? Date.new(2024, 1, 1) + i : Time.utc(2024, 1, 1) + i },
  'mixed' => words.first(4) + [42, 3.14, true, false, nil, Date.new(2024, 1, 1), Time.utc(2024, 1, 1), :symbol]
}.freeze

result = WORKLOADS.map do |name, values|
  ROWS.times { |i| Xlsxtream::Row.new(values, i + 1).to_xml } # warmup
  best = Array.new(ITERATIONS) do
    started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    ROWS.times { |i| Xlsxtream::Row.new(values, i + 1).to_xml }
    Process.clock_gettime(Process::CLOCK_MONOTONIC) - started
  end.min
  [name, best]
end.to_h

puts JSON.generate(result)
