#!/usr/bin/ruby
# $Id: farm.rbx,v 1.1 2004/05/09 01:57:09 aamine Exp $

$BitChannelFarm ||= false
unless $BitChannelContext
  load './farmrc'
  setup_environment
  require 'bitchannel/farm'
  require 'bitchannel/cgi'
  $BitChannelContext = farm_context()
end
BitChannel::FarmCGI.main $BitChannelContext
