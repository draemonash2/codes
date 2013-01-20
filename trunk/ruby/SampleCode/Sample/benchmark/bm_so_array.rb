#!/usr/bin/ruby
# -*- mode: ruby -*-
# $Id: bm_so_array.rb 24 2012-11-23 10:13:10Z TatsuyaEndo $
# http://www.bagley.org/~doug/shootout/
# with help from Paul Brannan and Mark Hubbart

n = 9000 # Integer(ARGV.shift || 1)

x = Array.new(n)
y = Array.new(n, 0)

n.times{|bi|
  x[bi] = bi + 1
}

(0 .. 999).each do |e|
  (n-1).step(0,-1) do |bi|
    y[bi] += x.at(bi)
  end
end
# puts "#{y.first} #{y.last}"


