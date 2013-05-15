#! ruby -Ks
# coding: windows-31j
# =================================================
#	$Brief	音楽ファイル編集を実施する $
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
#		なし
#	
# =================================================
# 
#! /usr/bin/env ruby

# =================================================
# Require 指定
# =================================================
require "./lib/input_file.rb"
require 'FileUtils'
require "find"

# =================================
# パラメータ設定
# =================================
EDIT_TARGET_PATH	= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
JACKET_DIR_PATH		= "D:/200_Pictures/30_CDJacket"

# =================================
# 実行処理
# =================================
# 事前確認処理
def check_preProcess()
	puts "以下が完了していることを確認してください。"
	puts "	01 Desktop に EditMP3Files フォルダ作成、 Mp3 移動"
	puts "	02 タグ名取得 @Amazon"
	puts "	03 タグ付け @SuperTagEditor"
	puts "	04 ジャケット取得 @Amazon"
	puts "	05 取得したジャケットを各フォルダに格納"
	STDIN.gets
	
	# JACKET_DIR_PATH  が存在するか確認
	# TODO : Implements this function
	
	# EDIT_TARGET_PATH が存在するか確認
	# TODO : Implements this function
	
	# EDIT_TARGET_PATH 配下のフォルダに 画像ファイルが格納されていることを確認
	# TODO : Implements this function
end

# 画像編集処理
def create_folderJPG()
	# AlbumArt.jpg と Folder.jpg を削除
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /AlbumArt_.*\.jpg/ ||
		   strFilePath =~ /Folder.*\.jpg/
			FileUtils.rm(strFilePath)
		end
	}
	
	# Folder.jpg にリネーム
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if File.extname(strFilePath) == ".jpg"
			File.rename( strFilePath, File.dirname(strFilePath) + "/Folder.jpg")
		end
	}
	
	# Folder.jpg 隠しファイル化
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if File.extname(strFilePath) == ".jpg"
			system("attrib \"#{strFilePath}\" +h")
		end
	}
end

# 画像ファイル退避処理
def copy_toJacketsDir()
	# JACKET_DIR_PATH から画像ファイルの最終ファイル名を取得
	fixMaxFileNumber	= 0
	Find.find(JACKET_DIR_PATH) {|strFilePath|
			if strFilePath =~ /Folder(...)\.jpg/
			if fixMaxFileNumber < $1.to_i
				fixMaxFileNumber = $1.to_i
			end
		end
	}
	
	# EDIT_TARGET_PATH 配下に格納された Folder.jpg を FolderXXX.jpg としてコピー
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

# MP3Gain 実行
def execute_MP3Gein()
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /.*\.mp3/
			system("mp3gain /r /c /p /d 14 \"#{strFilePath}\"") # 100db (86 + 14)
		end
	}
end

# Mp3tag 実行
def execute_MP3Tag()
	# TODO
end

# =================================
# 本処理
# =================================
begin
	# 事前確認処理
	check_preProcess()
	
	# 画像編集処理
	create_folderJPG()
	
	# 画像ファイル退避処理
	copy_toJacketsDir()
	
	# MP3Gain 実行
	execute_MP3Gein()
	
	# Mp3tag 実行
	execute_MP3Tag()
	
rescue
  # 例外処理
ensure
  # 後始末
end

