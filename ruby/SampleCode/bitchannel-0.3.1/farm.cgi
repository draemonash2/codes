#!/usr/bin/ruby
# $Id: farm.cgi,v 1.2 2004/05/09 01:57:09 aamine Exp $
#
# !!!! WARNING !!!!
# Never use this file for multisession environment e.g. mod_ruby, esehttpd.
# Use index.rbx instead.
#

load './farmrc'
setup_environment
require 'bitchannel/farm'
require 'bitchannel/cgi'
BitChannel::FarmCGI.main farm_context()
