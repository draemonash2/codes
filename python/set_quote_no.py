#!/usr/bin/env python3

# usage : python3 set_quote_no.py <file>

import re
import sys
import shutil
import os

def main():
    args = sys.argv
    if len(args) == 2:
        pass
    else:
        print('wrong number of arguments')
        return 0
    
    in_file_name = args[1]
    out_file_name = args[1] + ".tmp"
    
    pattern = r"^(.*)\[\[.*\]\](.*)"
    
    quote_idx = 1
    try:
        out_file = open(out_file_name, 'w')
        in_file = open(in_file_name)
        lines = in_file.readlines()
        for line in lines:
            matchlist = re.findall(pattern, line)
            if matchlist:
                out_file.write(matchlist[0][0] + "[[" + str(quote_idx) + "]]" + matchlist[0][1] + "\n")
                quote_idx += 1
            else:
                out_file.write(line)
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()
        shutil.copyfile(out_file_name, in_file_name)
        os.remove(out_file_name)

if __name__ == "__main__":
    main()

