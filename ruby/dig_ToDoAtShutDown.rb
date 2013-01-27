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
DESKTOP_PATH	= "#{ENV['USERPROFILE']}/Desktop".gsub("\\","/")

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
def push_automaticGithub(strDirPath, strPushDir)
	Dir::chdir(strDirPath)
	system("git add -u \.")					# 官吏対象ファイルのみコミット対象とする
	system("git commit -m \"Auto Commit\"")	# 変更のコミット
	system("git push #{strPushDir}")
end

# =================================
# 本処理
# =================================
	# RightsNetworkMediaPlugIn を削除する
	delete_RightsNetworkMediaPlugIn()

	# Setting コミット & プッシュ
	strDirPath	= "C:/prg"
	strPushDir	= "setting"
	push_automaticGithub(strDirPath, strPushDir)
