#!/usr/bin/ruby
# $Id: farm.fcgi,v 1.1 2004/05/09 01:57:09 aamine Exp $

load './farmrc'
setup_environment
require 'bitchannel/farm'
require 'bitchannel/fcgi'
BitChannel::FarmCGI.main farm_context()
