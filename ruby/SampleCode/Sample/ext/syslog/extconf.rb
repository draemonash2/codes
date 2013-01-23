# $RoughId: extconf.rb,v 1.3 2001/11/24 17:49:26 knu Exp $
# $Id: extconf.rb 24 2012-11-23 10:13:10Z  $

require 'mkmf'

have_header("syslog.h") &&
  have_func("openlog") &&
  have_func("setlogmask") &&
  create_makefile("syslog")

