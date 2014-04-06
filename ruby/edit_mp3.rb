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
EDIT_TARGET_PATH		= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
EDIT_TARGET_BAK_PATH	= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files_bak".gsub("\\", "/")
JACKET_DIR_PATH			= "Z:/200_Pictures/30_CDJacket"

# =================================
# ���s����
# =================================
# ���O����
def execute_preProcess()
	# ��ƃt�H���_�쐬
	FileUtils.mkdir_p EDIT_TARGET_PATH
	
	# �t�H���_���݊m�F
	if File.exists?(JACKET_DIR_PATH) == false then raise "Directry \"#{JACKET_DIR_PATH}\"is nothing!" end
	if File.exists?(EDIT_TARGET_PATH) == false then raise "Directry \"#{EDIT_TARGET_PATH}\"is nothing!" end
end

# �摜�ҏW����
def create_folderJPG()
	
	strJacketPath = Array.new()
	
	# �������\��
	puts " ********************************************* "
	puts " **          �W���P�b�g�����J�n!!           ** "
	puts " ********************************************* "
	puts "�ȉ����m�F���Ă��������B"
	puts "	1. �擾�����W���P�b�g���e�t�H���_�Ɋi�[����Ă��邱�ƁB "
	puts "	  ���e�t�H���_�ɂ̂݊i�[���Ă���������"
	STDIN.gets()
	
	# �W���P�b�g�p�X�擾
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /.*\.jpg/
			if strFilePath =~ /AlbumArt.*\.jpg/ ||
			   strFilePath =~ /Folder.*\.jpg/
				# AlbumArt.jpg �� Folder.jpg ���폜
				FileUtils.rm(strFilePath)
			else
				# �W���P�b�g�p�X�擾
				strJacketPath.push(strFilePath)
			end
		end
	}
	
	# �W���P�b�g�̃o�b�N�A�b�v
	FileUtils.mkdir_p(EDIT_TARGET_BAK_PATH)unless FileTest.exist?(EDIT_TARGET_BAK_PATH)
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		FileUtils.cp( strJacketPath[fixJacketPathCnt], EDIT_TARGET_BAK_PATH + "/" + File.basename(strJacketPath[fixJacketPathCnt]))
	end
	
	# Folder.jpg �Ƀ��l�[���{ �B���t�@�C����
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		File.rename( strJacketPath[fixJacketPathCnt], File.dirname(strJacketPath[fixJacketPathCnt]) + "/Folder.jpg")
		strJacketPath[fixJacketPathCnt].sub!( File.basename(strJacketPath[fixJacketPathCnt]), "Folder.jpg")
		system("attrib \"#{strJacketPath[fixJacketPathCnt]}\" +h")
		end
	
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
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		strSrcFilePath = strJacketPath[fixJacketPathCnt]
		strDstFilePath = JACKET_DIR_PATH + "/Folder#{format("%3d", fixMaxFileNumber)}.jpg"
		FileUtils.cp( strSrcFilePath, strDstFilePath)
		system("attrib \"#{strDstFilePath}\" -h")
		fixMaxFileNumber += 1
	end
	
	# �������\��
	puts " ********************************************* "
	puts " **          �W���P�b�g��������!!           ** "
	puts " ********************************************* "
	puts " �W���P�b�g�i�[�������������܂����̂ňȉ������{���Ă��������B"
	puts "	1. ���e�t�H���_�ɂ̂݊i�[�����W���P�b�g���e�f�B�X�N�̃t�H���_�Ɉړ����Ă��������B"
	STDIN.gets()
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

# ���㏈��
def execute_postProccess()
	puts "���삪�������܂����B�ȉ������{���Ă��������B"
	puts "	01 �^�O�t�� @SuperTagEditor"
	STDIN.gets()
end

# =================================
# �{����
# =================================
begin
	
	# ���O����
	execute_preProcess()
	
	# �摜�ҏW����
	create_folderJPG()
	
	# MP3Gain ���s
	execute_MP3Gein()
	
	# Mp3tag ���s
	execute_MP3Tag()
	
	# ���㏈��
	execute_postProccess()
	
rescue => error
	puts error.message
	puts error.backtrace
ensure
  # ��n��
end

