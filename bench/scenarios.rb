# frozen_string_literal: true
require 'date'

# Benchmark scenarios shared by bench/run.rb and bench/compare.rb.
#
# Each scenario names the worksheet options to use and how to build row i.
# Rows are built while writing, as a streaming caller would, and the cost of
# building them is measured separately.
#
# Options unknown to an older version of the gem are silently ignored by it,
# so a scenario for a new option compares "using the option" against "not
# having it", which is the cost of using the option.
module Bench
  COLUMNS = 10

  SCENARIOS = {
    'strings' => {
      :options => {},
      :row => ->(i) { Array.new(COLUMNS) { |c| "row #{i} col #{c}" } }
    },
    'strings_sst' => {
      :options => { :use_shared_strings => true },
      :row => ->(i) { Array.new(COLUMNS) { |c| "value #{(i * c) % 500}" } }
    },
    # Every string is unique, the worst case for the shared string table,
    # which keeps each distinct string in memory until the workbook closes.
    'strings_sst_unique' => {
      :options => { :use_shared_strings => true },
      :row => ->(i) { Array.new(COLUMNS) { |c| "row #{i} col #{c}" } }
    },
    'numbers' => {
      :options => {},
      :row => ->(i) { Array.new(COLUMNS) { |c| c.even? ? i * c : i * c / 7.0 } }
    },
    'dates_times' => {
      :options => {},
      :row => ->(i) {
        date = Date.new(2024, 1, 1) + (i % 365)
        time = Time.utc(2024, 1, 1) + i * 60
        Array.new(COLUMNS) { |c| c.even? ? date : time }
      }
    },
    'mixed' => {
      :options => { :has_header_row => true },
      :row => ->(i) {
        ["name #{i}", i, i / 3.0, Date.new(2024, 1, 1) + (i % 365), i.even?,
         nil, "x" * (i % 20), i * 100, Time.utc(2024, 1, 1) + i, 'fixed']
      }
    },
    'auto_format' => {
      :options => { :auto_format => true },
      :row => ->(i) { Array.new(COLUMNS) { |c| c.even? ? (i * c).to_s : "2024-01-#{(i % 28 + 1).to_s.rjust(2, '0')}" } }
    }
  }.freeze
end
