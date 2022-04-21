#!/bin/bash
TRGT=${1:-"cliant"}
gcc $TRGT.c -g -o $TRGT.out && ./$TRGT.out

