#!/usr/bin/ruby
# $Id: index.rbx,v 1.2 2004/05/09 01:43:01 aamine Exp $

load './bitchannelrc'
$BitChannelContext ||= nil
unless $BitChannelContext
  setup_environment
  $BitChannelContext = bitchannel_context()
end
require 'bitchannel/cgi'
BitChannel::CGI.main $BitChannelContext
