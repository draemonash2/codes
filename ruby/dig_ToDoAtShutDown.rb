#! /usr/bin/env ruby
# =================================================
#	$Brief	シャットダウン時に実行する処理を記述 $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author:  $
#	$HeadURL: file:///C:/Repo/trunk/ruby/dig_ToDoAtShutDown.rb $
#	
#	$UsageRule:
#		ruby dig_ToDoAtShutDown.rb
#	
#	$Note:
#		なし
#	
# =================================================

# =================================================
# Require 指定
# =================================================
require "find"

# =================================
# パラメータ設定
# =================================
DESKTOP_PATH			= "#{ENV['USERPROFILE']}/Desktop".gsub("\\","/")
GIT_MNG_DIR				= "C:/prg"
GIT_PUSH_REPO			= "setting"
HIDDEN_SETTING_FILE_DIR	= "C:/codes/vbs"

# =================================
# 実行処理
# =================================
# RightsNetworkMediaPlugIn を削除する
def delete_RightsNetworkMediaPlugIn()
	Find.find(DESKTOP_PATH) {|strFilePath|
		if strFilePath =~ /RightsNetworkMediaPlugIn.*\.exe/
			File.delete(strFilePath)
		end
	}
end

# コミット & プッシュ
def push_automaticGithub()
	Dir::chdir(GIT_MNG_DIR)
	system("git add -u \.")					# 監視対象ファイルのみコミット対象とする
	system("git commit -m \"Auto Commit\"") # 変更のコミット
	system("git push #{GIT_PUSH_REPO}")
end

# システムファイル、隠しファイルの非表示
def hide_systemFiles()
	Dir.chdir(HIDDEN_SETTING_FILE_DIR)
	system("HiddenSystemFiles.vbs")
end

# =================================
# 本処理
# =================================
#	# RightsNetworkMediaPlugIn を削除する
#	delete_RightsNetworkMediaPlugIn()
#
#	# Setting コミット & プッシュ
#	push_automaticGithub()
	
	# システムファイル、隠しファイルの非表示
	hide_systemFiles()
