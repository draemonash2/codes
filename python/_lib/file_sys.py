#!/usr/bin/env python3

import os

# Create object(file or directory) list.
#
#   arg1: root_dir_path     root directory path
#   arg2: object_type       object type as follows
#                               0:both, 1:files,
#                               2:directorys, other:none
#   return: object_list     object list
#   note:
#     - Only unix paths(/) can be specified.
#     - This function does not sort elements.
def create_file_list(root_dir_path, object_type = 0):
    object_list = []
    if not os.path.exists(root_dir_path):
        print("[error] create_file_list() root_dir_path does not exist.")
        return object_list
    
    for root, dirs, files in os.walk(root_dir_path):
        if object_type == 0 or object_type == 2:
            if not root == root_dir_path:
                path = root
                object_list.append(path)
        if object_type == 0 or object_type == 1:
            for file in files:
                path = os.path.join(root, file)
                object_list.append(path)
    return object_list

####################
### TEST PROGRAM ###
####################
def _test_create_file_list():
    trgt_dir_path = os.path.dirname(os.path.abspath(__file__)) + "/test_create_file_list"
    _create_blank_dir(trgt_dir_path)
    _create_blank_file(trgt_dir_path + "/test01.txt")
    _create_blank_file(trgt_dir_path + "/test02.txt")
    _create_blank_dir(trgt_dir_path + "/test03")
    _create_blank_dir(trgt_dir_path + "/test04")
    _create_blank_file(trgt_dir_path + "/.test05")
    
    print("*** test start ***")
    print(create_file_list(trgt_dir_path))
    print("")
    print(create_file_list(trgt_dir_path, 0))
    print("")
    print(create_file_list(trgt_dir_path, 1))
    print("")
    print(create_file_list(trgt_dir_path, 2))
    print("")
    print(create_file_list(trgt_dir_path, 3))
    print("")
    print(create_file_list(trgt_dir_path + "_"))
    print("")
    print("*** test finished ***")
    
    os.remove(trgt_dir_path + "/test01.txt")
    os.remove(trgt_dir_path + "/test02.txt")
    os.remove(trgt_dir_path + "/.test05")
    os.rmdir(trgt_dir_path + "/test03")
    os.rmdir(trgt_dir_path + "/test04")
    os.rmdir(trgt_dir_path)

def _create_blank_file(filepath):
    if not os.path.exists(filepath):
        f = open(filepath, 'w')
        f.write('')
        f.close()

def _create_blank_dir(dirpath):
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)

if __name__ == "__main__":
    _test_create_file_list()

