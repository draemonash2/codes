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
    
    pattern = r'(\[\[)(\d+)(\]\])'
    
    try:
        out_file = open(out_file_name, 'w')
        in_file = open(in_file_name)
        lines = in_file.readlines()
        for line in lines:
            #print(line, end="")
            matchlist = re.findall(pattern, line)
            if matchlist:
                #print(matchlist[0][0] + matchlist[0][1] + matchlist[0][2], end="")
                out_file.write(matchlist[0][0] + matchlist[0][1] + matchlist[0][2] + "\n")
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()

if __name__ == "__main__":
    main()

