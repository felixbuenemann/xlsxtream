# frozen_string_literal: true
module Xlsxtream
  class SharedStringTable < Hash
    def initialize
      @references = 0
      super { |hash, string| hash[string] = hash.size }
    end

    def [](string)
      @references += 1
      super
    end

    def references
      @references
    end
  end
end
