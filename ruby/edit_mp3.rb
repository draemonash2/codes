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
EDIT_TARGET_PATH		= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
EDIT_TARGET_BAK_PATH	= "#{ENV['USERPROFILE']}/Desktop/EditMP3Files_bak".gsub("\\", "/")
JACKET_DIR_PATH			= "Z:/200_Pictures/30_CDJacket"

# =================================
# 実行処理
# =================================
# 事前処理
def execute_preProcess()
	# 作業フォルダ作成
	FileUtils.mkdir_p EDIT_TARGET_PATH
	
	# フォルダ存在確認
	if File.exists?(JACKET_DIR_PATH) == false then raise "Directry \"#{JACKET_DIR_PATH}\"is nothing!" end
	if File.exists?(EDIT_TARGET_PATH) == false then raise "Directry \"#{EDIT_TARGET_PATH}\"is nothing!" end
end

# 画像編集処理
def create_folderJPG()
	
	strJacketPath = Array.new()
	
	# 説明文表示
	puts " ********************************************* "
	puts " **          ジャケット処理開始!!           ** "
	puts " ********************************************* "
	puts "以下を確認してください。"
	puts "	1. 取得したジャケットが各フォルダに格納されていること。 "
	puts "	  ★親フォルダにのみ格納してください★"
	STDIN.gets()
	
	# ジャケットパス取得
	Find.find(EDIT_TARGET_PATH) {|strFilePath|
		if strFilePath =~ /.*\.jpg/
			if strFilePath =~ /AlbumArt.*\.jpg/ ||
			   strFilePath =~ /Folder.*\.jpg/
				# AlbumArt.jpg と Folder.jpg を削除
				FileUtils.rm(strFilePath)
			else
				# ジャケットパス取得
				strJacketPath.push(strFilePath)
			end
		end
	}
	
	# ジャケットのバックアップ
	FileUtils.mkdir_p(EDIT_TARGET_BAK_PATH)unless FileTest.exist?(EDIT_TARGET_BAK_PATH)
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		FileUtils.cp( strJacketPath[fixJacketPathCnt], EDIT_TARGET_BAK_PATH + "/" + File.basename(strJacketPath[fixJacketPathCnt]))
	end
	
	# Folder.jpg にリネーム＋ 隠しファイル化
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		File.rename( strJacketPath[fixJacketPathCnt], File.dirname(strJacketPath[fixJacketPathCnt]) + "/Folder.jpg")
		strJacketPath[fixJacketPathCnt].sub!( File.basename(strJacketPath[fixJacketPathCnt]), "Folder.jpg")
		system("attrib \"#{strJacketPath[fixJacketPathCnt]}\" +h")
		end
	
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
	for fixJacketPathCnt in 0 .. (strJacketPath.length - 1)
		strSrcFilePath = strJacketPath[fixJacketPathCnt]
		strDstFilePath = JACKET_DIR_PATH + "/Folder#{format("%3d", fixMaxFileNumber)}.jpg"
		FileUtils.cp( strSrcFilePath, strDstFilePath)
		system("attrib \"#{strDstFilePath}\" -h")
		fixMaxFileNumber += 1
	end
	
	# 説明文表示
	puts " ********************************************* "
	puts " **          ジャケット処理完了!!           ** "
	puts " ********************************************* "
	puts " ジャケット格納処理が完了しましたので以下を実施してください。"
	puts "	1. ★親フォルダにのみ格納したジャケットを各ディスクのフォルダに移動してください。"
	STDIN.gets()
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

# 事後処理
def execute_postProccess()
	puts "操作が完了しました。以下を実施してください。"
	puts "	01 タグ付け @SuperTagEditor"
	STDIN.gets()
end

# =================================
# 本処理
# =================================
begin
	
	# 事前処理
	execute_preProcess()
	
	# 画像編集処理
	create_folderJPG()
	
	# MP3Gain 実行
	execute_MP3Gein()
	
	# Mp3tag 実行
	execute_MP3Tag()
	
	# 事後処理
	execute_postProccess()
	
rescue => error
	puts error.message
	puts error.backtrace
ensure
  # 後始末
end

