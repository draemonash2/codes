# $RoughId: extconf.rb,v 1.6 2001/07/13 15:38:27 knu Exp $
# $Id: extconf.rb 24 2012-11-23 10:13:10Z  $

require "mkmf"

$INSTALLFILES = {
  "digest.h" => "$(HDRDIR)"
}

create_makefile("digest")
