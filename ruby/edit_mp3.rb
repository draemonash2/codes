#! ruby -Ks
# coding: windows-31j
# =================================================
#	$Brief	���y�t�@�C���ҏW�����{���� $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author:  $
#	$HeadURL: file:///C:/Repo/trunk/ruby/edit_mp3.rb $
#	
#	$UsageRule:
#		ruby edit_mp3.rb
#	
#	$Note:
#		�Ȃ�
#	
# =================================================
# 
#! /usr/bin/env ruby

# =================================================
# Require �w��
# =================================================
require "./lib/input_file.rb"
require 'FileUtils'
require "find"

# =================================
# �p�����[�^�ݒ�
# =================================
EDIT_TARGET_PATH	= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
JACKET_DIR_PATH		= "D:/200_Pictures/30_CDJacket"

# =================================
# ���s����
# =================================
# ���O�m�F����
def check_preProcess()
	puts "�ȉ����������Ă��邱�Ƃ��m�F���Ă��������B"
	puts "	01 Desktop �� EditMP3Files �t�H���_�쐬�A Mp3 �ړ�"
	puts "	02 �^�O���擾 @Amazon"
	puts "	03 �^�O�t�� @SuperTagEditor"
	puts "	04 �W���P�b�g�擾 @Amazon"
	puts "	05 �擾�����W���P�b�g���e�t�H���_�Ɋi�["
	STDIN.gets
	
	# JACKET_DIR_PATH  �����݂��邩�m�F
	# TODO : Implements this function
	
	# EDIT_TARGET_PATH �����݂��邩�m�F
	# TODO : Implements this function
	
	# EDIT_TARGET_PATH �z���̃t�H���_�� �摜�t�@�C�����i�[����Ă��邱�Ƃ��m�F
	# TODO : Implements this function
end

# �摜�ҏW����
def create_folderJPG()
	# AlbumArt.jpg �� Folder.jpg ���폜
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /AlbumArt_.*\.jpg/ ||
		   strFilePath =~ /Folder.*\.jpg/
			FileUtils.rm(strFilePath)
		end
	}
	
	# Folder.jpg �Ƀ��l�[��
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if File.extname(strFilePath) == ".jpg"
			File.rename( strFilePath, File.dirname(strFilePath) + "/Folder.jpg")
		end
	}
	
	# Folder.jpg �B���t�@�C����
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if File.extname(strFilePath) == ".jpg"
			system("attrib \"#{strFilePath}\" +h")
		end
	}
end

# �摜�t�@�C���ޔ�����
def copy_toJacketsDir()
	# JACKET_DIR_PATH ����摜�t�@�C���̍ŏI�t�@�C�������擾
	fixMaxFileNumber	= 0
	Find.find(JACKET_DIR_PATH) {|strFilePath|
			if strFilePath =~ /Folder(...)\.jpg/
			if fixMaxFileNumber < $1.to_i
				fixMaxFileNumber = $1.to_i
			end
		end
	}
	
	# EDIT_TARGET_PATH �z���Ɋi�[���ꂽ Folder.jpg �� FolderXXX.jpg �Ƃ��ăR�s�[
	fixMaxFileNumber += 1
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /Folder.*\.jpg/
			strSrcFileName	= File.dirname(strFilePath) + "/" + "Folder.jpg"
			strDstFileName	= JACKET_DIR_PATH			+ "/" + "Folder.jpg"
			FileUtils.cp(strSrcFileName, strDstFileName)
			File.rename(strDstFileName, JACKET_DIR_PATH + "/Folder#{format("%3d", fixMaxFileNumber)}.jpg")
			system("attrib \"#{JACKET_DIR_PATH}/Folder#{format("%3d", fixMaxFileNumber)}.jpg\" -h")
			fixMaxFileNumber += 1
		end
	}
end

# MP3Gain ���s
def execute_MP3Gein()
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /.*\.mp3/
			system("mp3gain /r /c /p /d 14 \"#{strFilePath}\"") # 100db (86 + 14)
		end
	}
end

# Mp3tag ���s
def execute_MP3Tag()
	# TODO
end

# =================================
# �{����
# =================================
begin
	# ���O�m�F����
	check_preProcess()
	
	# �摜�ҏW����
	create_folderJPG()
	
	# �摜�t�@�C���ޔ�����
	copy_toJacketsDir()
	
	# MP3Gain ���s
	execute_MP3Gein()
	
	# Mp3tag ���s
	execute_MP3Tag()
	
rescue
  # ��O����
ensure
  # ��n��
end

