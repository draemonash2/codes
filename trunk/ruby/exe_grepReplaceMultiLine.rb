#! /usr/bin/env ruby
# =================================================
#	$Brief	複数行の置換を実施する $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/exe_grepReplaceMultiLine.rb $
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
require "./lib/input_file.rb"

# =================================================
# パラメータ指定
# =================================================

# =================================================
# 実行処理
# =================================================

# =================================================
# 本処理
# =================================================
strReplaceTargetDirPath	= ARGV[0].gsub("\\","/")
strReplaceFilePath		= ARGV[1].gsub("\\","/")
fixReplaceLine			= ARGV[2].to_i
matchSearchLine			= Regexp.new(ARGV[3])
strFileType				= ARGV[4]

# 置換対象ファイル一覧抽出
arrExtFiles	= Array.new()
if strFileType == "none"
	extract_file_path(strReplaceTargetDirPath, arrExtFiles, "none")
else
	extract_file_path(strReplaceTargetDirPath, arrExtFiles, strFileType)
end

# 置換行を入力
arrReplaceTxtFile	= Array.new()
input_txt(strReplaceFilePath, arrReplaceTxtFile)

# ターゲットフォルダ配下の一括置換
for fixFileCnt in 0 .. (arrExtFiles.length - 1)
	# 置換対象ファイルを入力
	arrReplaceTargetFile	= Array.new()
	input_txt(strReplaceTargetDirPath, arrReplaceTargetFile)
	
	# 置換
	replace_lines_byWord(arrReplaceTargetFile, fixReplaceLine, matchSearchLine, arrReplaceTxtFile)
end
