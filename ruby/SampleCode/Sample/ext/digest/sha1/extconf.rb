# $RoughId: extconf.rb,v 1.3 2001/08/14 19:54:51 knu Exp $
# $Id: extconf.rb 24 2012-11-23 10:13:10Z  $

require "mkmf"

$defs << "-DHAVE_CONFIG_H"
$INCFLAGS << " -I$(srcdir)/.."

$objs = [ "sha1init.#{$OBJEXT}" ]

dir_config("openssl")
pkg_config("openssl")
require_relative '../../openssl/deprecation'

if !with_config("bundled-sha1") &&
    have_library("crypto") && OpenSSL.check_func("SHA1_Transform", "openssl/sha.h")
  $objs << "sha1ossl.#{$OBJEXT}"
else
  $objs << "sha1.#{$OBJEXT}"
end

have_header("sys/cdefs.h")

$preload = %w[digest]

create_makefile("digest/sha1")
