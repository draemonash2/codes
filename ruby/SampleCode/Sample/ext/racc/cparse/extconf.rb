# $Id: extconf.rb 24 2012-11-23 10:13:10Z  $

require 'mkmf'
have_func('rb_block_call', 'ruby/ruby.h')
create_makefile 'racc/cparse'
