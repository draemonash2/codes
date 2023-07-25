#!/usr/bin/env python3

import sys
sys.path.append("/mnt/c/codes/python")

from _lib import file_sys
import os

filelist = file_sys.create_file_list("/mnt/c/codes/linux", 0)
print(filelist)

