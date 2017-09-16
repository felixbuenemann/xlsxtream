require 'test_helper'
require 'xlsxtream/shared_string_table'

module Xlsxtream
  class SharedStringTableTest < Minitest::Test
    def test_references
      sst = SharedStringTable.new
      assert_equal 0, sst.references
      sst['hello']
      sst['hello']
      sst['world']
      assert_equal 3, sst.references
    end

    def test_size
      sst = SharedStringTable.new
      assert_equal 0, sst.size
      sst['hello']
      sst['hello']
      sst['world']
      assert_equal 2, sst.size
    end

    def test_value
      sst = SharedStringTable.new
      assert_equal 0, sst['hello']
      assert_equal 1, sst['world']
      assert_equal 0, sst['hello']
    end
  end
end
