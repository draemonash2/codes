#!/bin/bash
TRGT=${1:-"cliant"}
g++ $TRGT.cpp -g -o $TRGT.out && ./$TRGT.out

