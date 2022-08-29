# -*- coding: cp932 -*-

import re
targetstr = r'123, 987, ,adb, 433, end.'
pattern = '((\d+), (\d+))'

print("targetstr = " + targetstr)

patobj = re.compile(pattern, re.I)
#matchlist = patobj.findall(targetstr)
##matchlist = re.findall(pattern, targetstr, re.I)
#print(type(matchlist))
#if matchlist:
#	print(len(matchlist))
#	print(matchlist[0])
#	print(len(matchlist[0]))
#	print(matchlist[0][0])
#	print(matchlist[0][2])
##	print(matchlist[1])

##matchobj = re.match(pattern, targetstr, re.I)
##matchobj = re.search(pattern, targetstr, re.I)
#matchobj = patobj.search(targetstr)
#print(type(matchobj))
#if matchobj:
##	print(len(matchobj))
#	print(matchobj.group())
#	print(matchobj.group(0))
#	print(matchobj.group(1))
#	print(matchobj.group(2))
#	print(matchobj.group(3))
#	print(matchobj.start(3))
#	print(matchobj.end(3))
	

repstr = re.sub(pattern, "xxx", targetstr)
print(repstr)

