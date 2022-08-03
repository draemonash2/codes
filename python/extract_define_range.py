#!/usr/bin/env python3

# usage : python3 extract_define_range.py <infile> <outfile> <define_keyword> <remain_target_side>
#    <remain_target_side>
#       if   : remain if side
#       else : remain else side

import re
import sys
import shutil
import os

def main():
    args = sys.argv
    if len(args) == 5:
        pass
    else:
        print('Arguments are too short')
        print('  usage : python3 extract_define_range.py <infile> <outfile> <define_keyword> <remain_target_side>')
        print('     <remain_target_side>')
        print('        if   : remain if side')
        print('        else : remain else side')
        return 0
    
    in_file_name = args[1]
    out_file_name = args[2]
    define_keyword = args[3]
    remain_target_side = args[4]
    assert remain_target_side == "if" or remain_target_side == "else"
    same_file_name = False
    if in_file_name == out_file_name:
        shutil.copyfile(in_file_name, in_file_name + '.tmp')
        in_file_name = in_file_name + '.tmp'
        same_file_name = True
    
    try:
        out_file = open(out_file_name, 'w')
        in_file = open(in_file_name)
        lines = in_file.readlines()
        is_remain = True
        for line in lines:
            matchresult_if     = re.match(r"^#if " + define_keyword + "$", line)
            matchresult_else   = re.match(r"^#else \/\* " + define_keyword + " \*\/$", line)
            matchresult_endif  = re.match(r"^#endif \/\* " + define_keyword + " \*\/$", line)
            matchresult_define = re.match(r"^#define " + define_keyword + " ", line)
            match_timing = False
            if matchresult_if:
                if remain_target_side == 'if':
                    is_remain = True
                else:
                    is_remain = False
                match_timing = True
            elif matchresult_else:
                if remain_target_side == 'if':
                    is_remain = False
                else:
                    is_remain = True
                match_timing = True
            elif matchresult_endif:
                is_remain = True
                match_timing = True
            elif matchresult_define:
                match_timing = True
            else:
                match_timing = False
            if match_timing == False and is_remain == True:
                out_file.write(line)
        if same_file_name == True:
            os.remove(in_file_name)
        print("extract finished to " + out_file_name)
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()

if __name__ == "__main__":
    main()

