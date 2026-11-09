# frozen_string_literal: true
# Times XML.escape_value for strings that do and don't need escaping and prints
# {"workload" => seconds} as JSON. Meant to be run through test/bench/compare.rb,
# but works on its own:
#
#   ruby -Ilib test/bench/escape_value_bench.rb
require 'json'
require 'xlsxtream/xml'

CALLS = Integer(ENV.fetch('CALLS', '500000'))
ITERATIONS = 3

WORKLOADS = {
  'short_plain' => 'foo_bar',
  'long_plain' => 'The quick brown fox jumps over the lazy dog, ' * 4,
  'utf8_plain' => 'Crème brûlée à la française – 日本語',
  'ampersand' => 'Smith & Sons <Ltd>',
  'hex_like' => 'ref_x1234_abc',
  'invalid_char' => "bell\x07ring"
}.freeze

result = WORKLOADS.map do |name, value|
  CALLS.times { Xlsxtream::XML.escape_value(value) } # warmup
  best = Array.new(ITERATIONS) do
    started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    CALLS.times { Xlsxtream::XML.escape_value(value) }
    Process.clock_gettime(Process::CLOCK_MONOTONIC) - started
  end.min
  [name, best]
end.to_h

puts JSON.generate(result)
