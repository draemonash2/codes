#!/bin/bash

function exectest() {
	echo "$ $@"
	"$@"
}

\cp -f test01.txt{.org.txt,}

# normal test
exectest python3 ../../set_quote_no.py test01.txt

# abnormal test
exectest python3 ../../set_quote_no.py
exectest python3 ../../set_quote_no.py test02.txt
exectest python3 ../../set_quote_no.py test01.txt aaa.txt

