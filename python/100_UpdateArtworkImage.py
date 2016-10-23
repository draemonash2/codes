# -*- coding: utf8 -*-

# <<Feature>>
#   - add folder.jpg to .mp3 file as artwork image
#   - if folder hasn't folder.jpg delete artwork image from .mp3 file
#   - hide folder.jpg
#   - [TODO] backup jpg image to temp backup folder
#   - [TODO] extract artwork image and save as jpg file
# 
# <<Usage>>
#   there is two way to execute, follow as:
#     1. double click me.
#     2. drag a target folder and drop to me.

import eyed3
import sys
import os

if len( sys.argv ) == 1:
	trgt_dir_path = "C:/Users/draem_000/Desktop/RippingMp3"
elif len( sys.argv ) == 2:
	trgt_dir_path = sys.argv[1].replace("\\", "/")
else:
	print "[error] argument is to many!"
	raw_input("\npress any key...")
	sys.exit()

log_file_path = ( trgt_dir_path + "/" + os.path.basename(__file__).replace(".py", ".log") ).replace("\\", "/")

if os.path.isdir( trgt_dir_path ) == True:
	log_file = open( log_file_path , "w")
else:
	print "[error] target directory is nothing!"
	raw_input("\npress any key...")
	sys.exit()

for root, dirs, files in os.walk( trgt_dir_path ):
	
	##########################################
	# check jpg files
	##########################################
	folder_jpg_exist = False
	other_jpg_exist = False
	for file_path in files:
		
		file_path = (root + "/" + file_path).replace("\\","/")
		
		root_path, file_ext = os.path.splitext( file_path )
		file_name = os.path.basename( file_path )
		
		if file_name == "Folder.jpg":
			folder_jpg_exist = True
			image_path = file_path
			# [TODO] backup jpg image
		else:
			if file_ext == ".jpg":
				other_jpg_exist = True
				# [TODO] backup jpg image
			elif file_ext == ".mp3":
				pass
				# [TODO] extract image to jpg
			else:
				pass
	
	log_file.write( "\n" )
	log_file.write( ( "[dir] " + root ).replace("\\","/") + "\n" )
	log_file.write( "[folder_jpg_exist] " + str( folder_jpg_exist ) + "\n" )
	log_file.write( "[other_jpg_exist] " + str( other_jpg_exist ) + "\n" )
	
	##########################################
	# delete and add artwork to mp3 file
	##########################################
	for file_path in files:
		file_path = (root + "/" + file_path).replace("\\","/")
		root_path, file_ext = os.path.splitext( file_path )
		
		log_file.write( "[file_path] " + file_path + "  [ext] " + file_ext + "\n" )
		
		if file_ext == ".mp3":
			audiofile = eyed3.core.load( file_path )
			
			#### delete artwork ####
			images_len = len( audiofile.tag.images )
			for i in range( images_len ):
				audiofile.tag.images.remove( u"" )
			
			#### add artwork ####
			if folder_jpg_exist == True:
				imagedata = open( image_path , "rb" ).read()
				audiofile.tag.images.set( 3, imagedata, "image/jpeg")
			else:
				pass
			
			#### save artwork ####
			audiofile.tag.save()
			
		elif file_ext == ".jpg":
			#### hide folder.jpg ####
			exe_cmd = "attrib +h \"" + file_path.replace("/","\\") + "\""
			os.system( exe_cmd )
			
		else:
			pass
	
log_file.close()

