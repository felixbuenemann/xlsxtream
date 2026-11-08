# frozen_string_literal: true
# Runs a benchmark file against two versions of lib/ and compares the timings.
#
#   ruby test/bench/compare.rb test/bench/row_dispatch_bench.rb [BASE] [TARGET]
#
# BASE and TARGET are git refs. BASE defaults to master, TARGET to the working
# tree (uncommitted changes included). Every round runs the benchmark in a fresh
# Ruby process per version, alternating the order to even out drift.
#
# The benchmark file must print a JSON object of {"workload name" => seconds}.
#
#   ROUNDS=7            processes per version (default 7)
#   ACTIVE_SUPPORT=1    load ActiveSupport before xlsxtream, as a Rails app would
require 'json'
require 'open3'
require 'rbconfig'
require 'tmpdir'

ROOT = File.expand_path('../..', __dir__)
bench_file = ARGV[0] or abort "Usage: ruby #{$0} BENCH_FILE [BASE] [TARGET]"
bench_file = File.expand_path(bench_file)
base_ref = ARGV[1] || 'master'
target_ref = ARGV[2]
rounds = Integer(ENV.fetch('ROUNDS', '7'))

def lib_for(ref, dir)
  return File.join(ROOT, 'lib') if ref.nil?

  target = File.join(dir, ref.gsub(/[^\w.-]/, '_'))
  Dir.mkdir(target)
  archive, status = Open3.capture2('git', '-C', ROOT, 'archive', ref, 'lib', :binmode => true)
  abort "git archive #{ref} failed" unless status.success?
  _, status = Open3.capture2('tar', '-x', '-C', target, :stdin_data => archive, :binmode => true)
  abort "Extracting #{ref} failed" unless status.success?
  File.join(target, 'lib')
end

def run(lib, bench_file)
  # Bundler would put the working tree's lib/ on the load path, so run without it
  env = ENV.to_h.reject { |key, _| key.start_with?('BUNDLE', 'RUBYOPT', 'RUBYLIB') }
  output, status = Open3.capture2(env, RbConfig.ruby, '-I', lib, bench_file, :unsetenv_others => true)
  abort "Benchmark failed with #{lib}" unless status.success?
  JSON.parse(output)
end

def median(values)
  sorted = values.sort
  mid = sorted.size / 2
  sorted.size.odd? ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2.0
end

def spread(values)
  (values.max - values.min) / median(values) * 100
end

Dir.mktmpdir('xlsxtream-bench') do |dir|
  libs = { :base => lib_for(base_ref, dir), :target => lib_for(target_ref, dir) }
  times = { :base => Hash.new { |h, k| h[k] = [] }, :target => Hash.new { |h, k| h[k] = [] } }

  rounds.times do |round|
    order = round.even? ? [:base, :target] : [:target, :base]
    order.each do |version|
      run(libs[version], bench_file).each { |name, seconds| times[version][name] << seconds }
    end
  end

  puts "#{File.basename(bench_file)}: #{base_ref} (base) vs #{target_ref || 'working tree'} (target)"
  puts "Ruby #{RUBY_VERSION}, #{rounds} rounds#{', ActiveSupport loaded' if ENV['ACTIVE_SUPPORT']}"
  puts
  puts format('%-24s %10s %10s %8s %7s', 'workload', 'base ms', 'target ms', 'change', 'noise')
  times[:base].each_key do |name|
    base = median(times[:base][name])
    target = median(times[:target][name])
    noise = [spread(times[:base][name]), spread(times[:target][name])].max
    change = (target - base) / base * 100
    flag = change.abs > noise ? '*' : ' '
    puts format('%-24s %10.1f %10.1f %+7.1f%%%s %5.1f%%', name, base * 1000, target * 1000, change, flag, noise)
  end
  puts
  puts 'change: target vs base, negative is faster; * marks a change larger than the noise'
  puts 'noise: largest (max - min) / median across rounds of either version'
end
