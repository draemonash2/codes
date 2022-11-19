#!/bin/sh

# normal test
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.dbg-if.c DBG if
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.dbg-else.c DBG else
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modif-if.c MOD_IF if
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modif-else.c MOD_IF else
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifdef-if.c MOD_IFDEF if
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifdef-else.c MOD_IFDEF else
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifndef-if.c MOD_IFNDEF if
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.modifndef-else.c MOD_IFNDEF else

\cp -f cliant.c.org.c cliant.c
python3 ../../extract_define_range.py cliant.c cliant.c MOD_IF else

\cp -f cliant.c.org.c cliant.c

# abnormal test
python3 ../../extract_define_range.py cliant.c.org.c DBG if
python3 ../../extract_define_range.py cliant.c.or.c cliant.c.or.c DBG if
python3 ../../extract_define_range.py cliant.c.org.c cliant.c.org.c DBG ifs

python3 ../../extract_define_range.py cliant.c.org.c cliant.c.org.c DBGS if

