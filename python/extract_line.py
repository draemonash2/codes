#!/usr/bin/env python3

# usage : python3 extract_line.py <infile> <outfile>

import re
import sys

def main():
    args = sys.argv
    if len(args) == 3:
        pass
    else:
        print('Arguments are too short')
        return 0
    
    in_file_name = args[1]
    out_file_name = args[2]
    
    pattern = r'^(header:|    stamp:|        sec:|        nanosec:)'
    
    try:
        out_file = open(out_file_name, 'w')
        in_file = open(in_file_name)
        lines = in_file.readlines()
        for line in lines:
            result = re.match(pattern, line)
            if result: # except for none
                out_file.write(line)
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()

if __name__ == "__main__":
    main()

