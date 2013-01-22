#!/usr/bin/ruby
# $Id: index.fcgi,v 1.5 2004/05/09 01:43:01 aamine Exp $

load './bitchannelrc'
setup_environment
require 'bitchannel/fcgi'
BitChannel::FCGI.main bitchannel_context()
