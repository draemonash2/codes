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
#		ruby exe_grepReplaceMultiLine.rb <replace_target_dir_path> <replace_line_num> <search_word> <replace_file_path>
#			<replace_target_dir_path>	: 置換対象ディレクトリパス
#			<replace_line_num>			: 置換行数
#			<search_word>				: 検索文字列
#			<replace_file_path>			: 置換行ファイルパス
#			<replace_file_type>			: 置換対象ファイル種別
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
#	git add # 設定ファイルはファイルが増えることがないため add しない
#	TODO カレントディレクトリ変更処理追加
	system("git commit -m \"自動コミット\"")
	system("git push #{strPushDir}")
end

# =================================
# 本処理
# =================================
# RightsNetworkMediaPlugIn を削除する
delete_RightsNetworkMediaPlugIn()

# Vim Setting コミット & プッシュ
push_automaticGithub(strDirPath, strPushDir)

# X-Finder Setting コミット & プッシュ
push_automaticGithub(strDirPath, strPushDir)
