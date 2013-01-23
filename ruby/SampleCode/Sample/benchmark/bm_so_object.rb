#!/usr/bin/ruby
# -*- mode: ruby -*-
# $Id: bm_so_object.rb 24 2012-11-23 10:13:10Z  $
# http://www.bagley.org/~doug/shootout/
# with help from Aristarkh Zagorodnikov

class Toggle
    def initialize(start_state)
        @bool = start_state
    end

    def value
        @bool
    end

    def activate
        @bool = !@bool
        self
    end
end

class NthToggle < Toggle
    def initialize(start_state, max_counter)
        super start_state
        @count_max = max_counter
        @counter = 0
    end

    def activate
        @counter += 1
        if @counter >= @count_max
            @bool = !@bool
            @counter = 0
        end
        self
    end
end

n = 1500000 # (ARGV.shift || 1).to_i

toggle = Toggle.new 1
5.times do
    toggle.activate.value ? 'true' : 'false'
end
n.times do
    toggle = Toggle.new 1
end

ntoggle = NthToggle.new 1, 3
8.times do
    ntoggle.activate.value ? 'true' : 'false'
end
n.times do
    ntoggle = NthToggle.new 1, 3
end

