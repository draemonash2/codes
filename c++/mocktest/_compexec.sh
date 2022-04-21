#!/bin/bash
TRGT=${1:-"cliant"}
gcc -c mystring.c -o mystring.o
gcc $TRGT.c mystring.o -g -o $TRGT.out && ./$TRGT.out

