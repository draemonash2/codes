/* $Id: sha1ossl.c 24 2012-11-23 10:13:10Z  $ */

#include "defs.h"
#include "sha1ossl.h"

void
SHA1_Finish(SHA1_CTX *ctx, char *buf)
{
	SHA1_Final((unsigned char *)buf, ctx);
}
