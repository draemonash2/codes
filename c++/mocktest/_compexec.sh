#!/bin/bash
TRGT=${1:-"cliant"}
g++ $TRGT.cpp -o $TRGT.out && ./$TRGT.out

