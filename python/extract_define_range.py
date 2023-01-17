#!/usr/bin/env python3

# extract_define_range.py ver1.2
#
# usage : python3 extract_define_range.py <infile> <outfile> <define_keyword> <remain_target_side>
#    <remain_target_side>
#       true  : remain "true" side
#       false : remain "false" side
#         e.g. specified "true" side
#           #if AAA           #del
#              true side      #remain
#           #else /* AAA */   #del
#              false side     #del
#           #endif /* AAA */  #del
#
#           #ifdef AAA        #del
#              true side      #remain
#           #else /* AAA */   #del
#              false side     #del
#           #endif /* AAA */  #del
#
#           #ifndef AAA       #del
#              true side      #del
#           #else /* !AAA */  #del
#              false side     #remain
#           #endif /* !AAA */ #del

import re
import sys
import shutil
import os

def main():
    delete_org_file = True
    args = sys.argv
    if len(args) != 5:
        print('[error  ] arguments are too short')
        print('  usage : python3 extract_define_range.py <infile> <outfile> <define_keyword> <remain_target_side>')
        print('     <remain_target_side>')
        print('        true  : remain "true" side')
        print('        false : remain "false" side')
        return 0
    in_file_name = args[1]
    out_file_name = args[2]
    define_keyword = args[3]
    remain_target_side = args[4]
    if not os.path.exists(in_file_name):
        print('[error  ] ' + in_file_name + ' does not exist.')
        return 0
    if remain_target_side != "true" and remain_target_side != "false":
        print('[error  ] specified <remain_target_side> is \"' + remain_target_side + '\". this is \"true\" or \"false\" only.')
        return 0
    
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
        remove_executed = False
        iftype = ""
        linenum = 1
        for line in lines:
            matchresult_define = re.match(r"^#define " + define_keyword + " ", line)
            matchresult_if     = re.match(r"^#if " + define_keyword + "$", line)
            matchresult_ifdef  = re.match(r"^#ifdef " + define_keyword + "$", line)
            matchresult_ifndef = re.match(r"^#ifndef " + define_keyword + "$", line)
            matchresult_else   = re.match(r"^#else \/\* " + define_keyword + " \*\/$", line)
            matchresult_elsen  = re.match(r"^#else \/\* !" + define_keyword + " \*\/$", line)
            matchresult_endif  = re.match(r"^#endif \/\* " + define_keyword + " \*\/$", line)
            matchresult_endifn = re.match(r"^#endif \/\* !" + define_keyword + " \*\/$", line)
            
            match_timing = False
            if matchresult_define:
                match_timing = True
                remove_executed = True
            elif matchresult_if:
                if remain_target_side == 'true':
                    is_remain = True
                else:
                    is_remain = False
                match_timing = True
                remove_executed = True
                iftype = "IF"
            elif matchresult_ifdef:
                if remain_target_side == 'true':
                    is_remain = True
                else:
                    is_remain = False
                match_timing = True
                remove_executed = True
                iftype = "IFDEF"
            elif matchresult_ifndef:
                if remain_target_side == 'true':
                    is_remain = False
                else:
                    is_remain = True
                match_timing = True
                remove_executed = True
                iftype = "IFNDEF"
            elif matchresult_else:
                if iftype == "IF" or iftype == "IFDEF":
                    pass
                else:
                    print('[error  ] this "#else" must be preceded by an "#if" or "#ifdef" at line:' + linenum + '.')
                    return 0
                
                if remain_target_side == 'true':
                    is_remain = False
                else:
                    is_remain = True
                match_timing = True
                remove_executed = True
                iftype = ""
            elif matchresult_elsen:
                if iftype == "IFNDEF":
                    pass
                else:
                    print('[error  ] this "#else" must be preceded by an "#ifndef" at line:' + linenum + '.')
                    return 0
                
                if remain_target_side == 'true':
                    is_remain = True
                else:
                    is_remain = False
                match_timing = True
                remove_executed = True
                iftype = ""
            elif matchresult_endif:
                is_remain = True
                match_timing = True
                remove_executed = True
                iftype = ""
            elif matchresult_endifn:
                is_remain = True
                match_timing = True
                remove_executed = True
                iftype = ""
            else:
                match_timing = False
            
            if match_timing == False and is_remain == True:
                out_file.write(line)
            
            linenum = linenum + 1
        if remove_executed == True:
            result_str = "[success]"
        else:
            result_str = "[skip   ]"
        if same_file_name == True:
            if delete_org_file == True:
                os.remove(in_file_name)
            print(result_str + " \"" + remain_target_side + "\" side of \"" + define_keyword + "\" : " + out_file_name)
        else:
            print(result_str + " \"" + remain_target_side + "\" side of \"" + define_keyword + "\" : " + in_file_name + " -> " + out_file_name)
    except Exception as e:
        print(e)
    finally:
        out_file.close()
        in_file.close()

if __name__ == "__main__":
    main()

