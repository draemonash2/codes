#!/bin/bash

function exectest() {
	echo "$ $@"
	"$@"
}

# normal test
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.dbg-true.c DBG true
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.dbg-false.c DBG false
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modif-true.c MOD_IF true
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modif-false.c MOD_IF false
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifdef-true.c MOD_IFDEF true
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifdef-false.c MOD_IFDEF false
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifndef-true.c MOD_IFNDEF true
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifndef-false.c MOD_IFNDEF false

\cp -f cliant.c.org.c cliant.c
exectest python3 ../../extract_define_range.py cliant.c cliant.c MOD_IF false

\cp -f cliant.c.org.c cliant.c

# abnormal test
exectest python3 ../../extract_define_range.py cliant.c.org.c DBG true
exectest python3 ../../extract_define_range.py cliant.c.or.c cliant.c.or.c DBG true
exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.org.c DBG trues

exectest python3 ../../extract_define_range.py cliant.c.org.c cliant.c.org.c DBGS true

