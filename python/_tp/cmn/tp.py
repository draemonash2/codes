#!/usr/bin/env python3

#from datetime import datetime
#
#cur_datetime = datetime.now()
#
#print(cur_datetime.strftime("%Y/%m/%d %H:%M:%S"))
#
#
#prev_datetime = datetime(2022, 12, 31, 23, 12, 59)
#
#print((cur_datetime - prev_datetime).days)
##print((cur_datetime - prev_datetime).hours)
#print((cur_datetime - prev_datetime).seconds)
##print((cur_datetime - prev_datetime).milliseconds)
#print((cur_datetime - prev_datetime).microseconds)
#print((cur_datetime - prev_datetime).total_seconds())

fruits = {"apple": "りんご", "orange": "みかん", "peach": "もも"}
fruits["grape"] = "ぶどう"
fruits["apple"] = "あおりんご"

#fruits2={"orange":"いよかん", "banana":"バナナ"}
#fruits.update(fruits2) #キーがかぶっている場合は値を上書き
#print(fruits)

print(list(fruits.values()))

#del fruits["orange"]
#print(fruits)
#print(len(fruits))
#if "apple" in fruits:
#	print("exists")
#else:
#	print("do not exists")

#for fruit in fruits:
#	print(fruit)

#for fruit_english in fruits.keys():
#	print(fruit_english)
#
#for fruit_japanese in fruits.values():
#	print(fruit_japanese)
#
#for fruit_english, fruit_japanese in fruits.items():
#	print(fruit_english + ":" + fruit_japanese)
#fruits.clear()
#removed_value = fruits.pop('peach')
#
#print(fruits)
