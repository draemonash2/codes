#!/bin/bash


if [ $# -ne 3 ]; then
	echo "[error] arguments error."
	echo "  usage : extract_defines.sh <dirpath> <extract_keyword> <remain_side>"
	return 1
fi

dirpath=$1
extract_keyword=$2
remain_side=$3

filelist=`find ${dirpath} -type f`
for file in ${filelist}
do
	#echo ${file}
	python3 /mnt/c/codes/python/extract_define_range.py ${file} ${file} ${extract_keyword} ${remain_side}
done

