#!/usr/bin/ruby
# $Id: index.cgi,v 1.11 2004/05/09 01:43:01 aamine Exp $
#
# !!!! WARNING !!!!
# Never use this file for multisession environment e.g. mod_ruby, esehttpd.
# Use index.rbx instead.
#

load './bitchannelrc'
setup_environment
require 'bitchannel/cgi'
BitChannel::CGI.main bitchannel_context()
