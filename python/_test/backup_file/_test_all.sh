#!/bin/bash

echo "" > _test_all.log
filelist=$(find . -maxdepth 1 -type f | grep ./_test_backup_file | awk -F/ '{ print $2 }')

for file in $filelist
do
	./${file} >> _test_all.log
done

