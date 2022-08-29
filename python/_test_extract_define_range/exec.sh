#!/bin/sh

# normal test
python3 ../extract_define_range.py cliant.c.org.c cliant.c.dbgif.c DBG if
python3 ../extract_define_range.py cliant.c.org.c cliant.c.dbgelse.c DBG else
python3 ../extract_define_range.py cliant.c.org.c cliant.c.mod01if.c MOD01 if
python3 ../extract_define_range.py cliant.c.org.c cliant.c.mod01else.c MOD01 else
python3 ../extract_define_range.py cliant.c.org.c cliant.c.mod02if.c MOD02 if
python3 ../extract_define_range.py cliant.c.org.c cliant.c.mod02else.c MOD02 else

\cp -f cliant.c.org.c cliant.c
python3 ../extract_define_range.py cliant.c cliant.c MOD02 else

# abnormal test
python3 ../extract_define_range.py cliant.c.org.c DBG if
python3 ../extract_define_range.py cliant.c.or.c cliant.c.or.c DBG if
python3 ../extract_define_range.py cliant.c.org.c cliant.c.org.c DBG ifs

