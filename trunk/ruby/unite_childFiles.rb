#! /usr/bin/env ruby
# =================================================
#	$Brief	指定ディレクトリ配下のファイルを結合して、
#			ファイルに出力する
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/unite_childFiles.rb $
#	
#	$UsageRule:
#		ruby unite_childFiles.rb <dir_path> <file_type>
#		  dir_path	 : 結合対象ディレクトリパス
#		  file_type	 : 結合対象ファイル種別
#	
#	$Note:
#		・結合結果は ./unite_childFiles.txt に出力されます
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
	# ===============================================================
	# @brief	引数チェック処理
	#
	# @param	なし
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def check_type()
		if ARGV.length != 2
			puts "parameter length error!"
			STDIN.gets()
			exit
		end
		
		if ARGV[0] == ""
			puts "dir path error!"
			STDIN.gets()
			exit
		end

		if ARGV[1][0] != "."
			puts "error filetype!"
			puts " \".\" none!"
			STDIN.gets()
			exit
		end
	end
	
	# ===============================================================
	# @brief	ファイル内容を抽出する
	#
	# @param	arrExtFileInfo		[in]	Array->Array->String	入力ファイル情報
	# @param	arrExtFileContents	[out]	Array->Array->String	ファイル内容格納先
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def extract_file_contents(arrExtFileInfo, arrExtFileContents)
		for fixFileCnt in 0 .. (arrExtFileInfo.length - 1)
			fileInputFile = $stdin
			fileInputFile = File.open(arrExtFileInfo[fixFileCnt][0], "r")
			arrInputFile  = Array.new()
			arrInputFile  = fileInputFile.readlines
			fileInputFile.close
			
			arrExtFileContents.push(arrInputFile)
		end
	end
	
	# ===============================================================
	# @brief	ファイルを結合する
	#
	# @param	arrExtFileInfo		[in]	Array->Array->String	入力ファイル情報
	# @param	arrExtFileContents	[in]	Array->Array->String	入力ファイル内容
	# @param	arrUniteFile		[out]	Array->String			結合結果格納先
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def unite_files(arrExtFileInfo, arrExtFileContents, arrUniteFile)
		# パラメータチェック
		if arrExtFileInfo.length != arrExtFileContents.length
			puts "parameter length error!"
			puts __FILE__
			puts __LINE__
			STDIN.gets()
			exit
		end
		
		for fixFileCnt in 0 .. (arrExtFileInfo.length - 1)
			# ファイル情報
			arrUniteFile.push(format("file_size    = %30s", arrExtFileInfo[fixFileCnt][1])) # ファイルサイズ
			arrUniteFile.push(format("acc_time     = %30s", arrExtFileInfo[fixFileCnt][2])) # 最終アクセス時刻
			arrUniteFile.push(format("update_time  = %30s", arrExtFileInfo[fixFileCnt][3])) # 最終更新時刻
			arrUniteFile.push(format("file_type    = %30s", arrExtFileInfo[fixFileCnt][4])) # ファイルタイプ
			# ファイル内容
			for fixLineCnt in 0 .. (arrExtFileContents[fixFileCnt].length - 1)
				arrUniteFile.push(arrExtFileContents[fixFileCnt][fixLineCnt])
			end
		end
	end

# =================================================
# 本処理
# =================================================

strDirPath	= ARGV[0].gsub("\\", "/")
strFileType = ARGV[1]

arrExtFileInfo		= Array.new() # ファイル情報
arrExtFileContents	= Array.new() # ファイル内容
arrUniteFile		= Array.new() # 結合結果

strTgtPath	= File.basename(__FILE__) + ".txt"

# 引数チェック
check_type()

# ファイル情報 (ファイルパス ファイルサイズ 最終アクセス時 最終更新時刻 ファイルタイプ) 抽出
extract_file_info(strDirPath, arrExtFileInfo, strFileType)

# ファイル内容抽出
extract_file_contents(arrExtFileInfo, arrExtFileContents)

# ファイル結合
unite_files(arrExtFileInfo, arrExtFileContents, arrUniteFile)

# ファイル出力
output_txt(arrOutputArr, strTgtPath)
