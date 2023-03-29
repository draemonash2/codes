#!/usr/bin/env python3

# usage : python3 set_quote_no.py <file>

import re
import sys
import shutil
import os

def main():
    args = sys.argv
    if len(args) != 2:
        print('[error] wrong number of arguments')
        print('  usage : python3 set_quote_no.py <file>')
        return 0
    
    in_file_name = args[1]
    out_file_name = args[1] + ".tmp"
    
    if os.path.exists(in_file_name) == False:
        print('[error] file does not exist : ' + in_file_name)
        return 0
    
    pattern = r"(\[\[\d+\]\])"
    
    quote_idx = 1
    try:
        out_file = open(out_file_name, 'w')
        in_file = open(in_file_name)
        lines = in_file.readlines()
        for line in lines:
            matchlist = list(re.finditer(pattern, line))
            list_num = len(matchlist)
            list_idx = list_num
            for matchobj in reversed(matchlist):
                match_start_pos = matchobj.span()[0]
                match_end_pos = matchobj.span()[1]
                output_quote_idx = quote_idx + list_idx - 1
                line = line[:match_start_pos] + "[[" + str(output_quote_idx) + "]]" + line[match_end_pos:]
                list_idx -= 1
            #print(line)
            out_file.write(line)
            quote_idx += list_num
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()
        shutil.copyfile(out_file_name, in_file_name)
        os.remove(out_file_name)

if __name__ == "__main__":
    main()

