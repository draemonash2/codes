#!/bin/bash

# Script file path setting
filename=$(basename $0)
filepath=$(cd $(dirname $0); pwd)/${filename}
filedirpath=$(dirname ${filepath})
filebasename=${filename%%.*}
logfilepath=${filedirpath}/${filebasename}.log

# Log output target setting
exec 1> >(tee -a ${logfilepath})
exec 2> >(tee -a ${logfilepath})

# Output pre-messages
echo ""
echo "### backup start! ###"
echo "$(date)"

# Backup process
(
	cd ~/_dotfiles
	git add *
	git commit -m 'scheduled backup.'
)
