/* $Id: rmd160ossl.c 24 2012-11-23 10:13:10Z  $ */

#include "defs.h"
#include "rmd160ossl.h"

void RMD160_Finish(RMD160_CTX *ctx, char *buf) {
	RIPEMD160_Final((unsigned char *)buf, ctx);
}
