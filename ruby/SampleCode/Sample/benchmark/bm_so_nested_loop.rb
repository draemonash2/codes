#!/usr/bin/ruby
# -*- mode: ruby -*-
# $Id: bm_so_nested_loop.rb 24 2012-11-23 10:13:10Z  $
# http://www.bagley.org/~doug/shootout/
# from Avi Bryant

n = 16 # Integer(ARGV.shift || 1)
x = 0
n.times do
    n.times do
        n.times do
            n.times do
                n.times do
                    n.times do
                        x += 1
                    end
                end
            end
        end
    end
end
# puts x


