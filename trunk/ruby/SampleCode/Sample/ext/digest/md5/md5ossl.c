/* $Id: md5ossl.c 24 2012-11-23 10:13:10Z TatsuyaEndo $ */

#include "md5ossl.h"

void
MD5_Finish(MD5_CTX *pctx, unsigned char *digest)
{
    MD5_Final(digest, pctx);
}
