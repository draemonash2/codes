#!/bin/bash

CURDIR=`cd $(dirname ${0}) && pwd`
IN_FNAME=${CURDIR}/../.vimrc
OUT_FNAME_PRE=${CURDIR}/../_vimrc
OUT_FNAME_POST=${CURDIR}/../_gvimrc
KEYWORD='キーバインド設定'

keywordlineno=`grep -n ${KEYWORD} ${IN_FNAME} | cut -d: -f 1`
splitlineno=`expr $keywordlineno - 2`
cat ${IN_FNAME} | awk "NR<=$splitlineno {print}" > ${OUT_FNAME_PRE}
cat ${IN_FNAME} | awk "NR>$splitlineno {print}" > ${OUT_FNAME_POST}

