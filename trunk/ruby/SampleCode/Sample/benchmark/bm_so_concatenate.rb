#!/usr/bin/ruby
# -*- mode: ruby -*-
# $Id: bm_so_concatenate.rb 24 2012-11-23 10:13:10Z TatsuyaEndo $
# http://www.bagley.org/~doug/shootout/
# based on code from Aristarkh A Zagorodnikov and Dat Nguyen

STUFF = "hello\n"
i = 0
while i<10
  i += 1
  hello = ''
  4_000_000.times do |e|
    hello << STUFF
  end
end
# puts hello.length


