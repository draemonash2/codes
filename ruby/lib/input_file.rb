#! /usr/bin/env ruby
# =================================================
#	$Brief	ファイルの入力を行う $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/input_file.rb $
#	
# =================================================

require "find"

	# ===============================================================
	# @brief	CSVファイルを入力し、二次元配列に格納する
	#
	# @param	strTgtPath		[in]	String					入力ファイルパス
	# @param	arrInputArray	[in]	Array->Array->String	入力データ配列
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def input_csv(strTgtPath, arrInputArray)
		input_file = $stdin
		input_file = File.open(strTgtPath, 'r')
		input = Array.new
		input = input_file.readlines
		
		for i in 0..input.length - 1
			arrInputArray << input[i].split(",")
		end
	end

	# ===============================================================
	# @brief	TSVファイルを入力し、二次元配列に格納する
	#
	# @param	strTgtPath		[in]	String					入力ファイルパス
	# @param	arrInputArray	[in]	Array->Array->String	入力データ配列
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def input_tsv(strTgtPath)
		input_file = $stdin
		input_file = File.open(strTgtPath, 'r')
		input = Array.new
		input = input_file.readlines
		arrInputArray = Array.new
		
		for i in 0..input.length - 1
			arrInputArray << input[i].split("\t")
		end
		
		return arrInputArray
	end

	# ===============================================================
	# @brief	TXT ファイルを入力し、配列に格納する
	#
	# @param	strTargetPath	[in]	String			入力ファイルパス
	# @param	arrInputArray	[out]	Array->String	入力データ配列
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def input_txt(strTargetPath, arrInputArray)
		input_file = $stdin
		input_file = File.open(strTargetPath, 'r')
		input = Array.new
		input = input_file.readlines
		
		for i in 0..input.length - 1
			arrInputArray << input[i].split(",")
		end
		input_file.close
	end
	
	# ===============================================================
	# @brief	ファイルを抽出する
	#			ディレクトリは抽出しない
	#
	# @param	strTargetPath	[in]	String			入力ファイルパス
	# @param	arrExtFiles		[out]	Array->String	入力データ配列
	# @param	strFileType		[in]	String			入力ファイルタイプ
	#				none	: ファイル種別指定なし
	#				.xx		: ファイル種別指定あり( xx => 拡張子名)
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def extract_file_path(strTargetPath, arrExtFiles, strFileType)
	#	if strFileType[0]	!= "."
	#		puts "file type error!"
	#		puts __FILE__
	#		puts __LINE__
	#		STDIN.gets()
	#		exit
	#	end
		
		Find.find(strTargetPath) {|filepath|
			if File.ftype(filepath) == "file"
				if strFileType == "none"
					arrExtFiles.push(filepath)
				else
					if File.extname(filepath) == strFileType
						arrExtFiles.push(filepath)
					end
				end
			end
		}
	end
	
	# ===============================================================
	# @brief	ファイルを抽出する
	#			ディレクトリは抽出しない
	#			同時にファイル情報を抽出する
	#				・ファイルパス
	#				・ファイルサイズ
	#				・最終アクセス時刻
	#				・最終更新時刻
	#				・ファイルタイプ
	#
	# @param	strTargetPath	[in]	String					入力ファイルパス
	# @param	arrExtFiles		[out]	Array->Array->String	入力データ配列
	# @param	strFileType		[in]	String					入力ファイルタイプ
	#				none	: ファイル種別指定なし
	#				.xx		: ファイル種別指定あり( xx => 拡張子名)
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def extract_file_info(strTargetPath, arrExtFiles, strFileType)
		
		# ファイルパス 抽出
		extract_file_path(strTargetPath, arrExtFiles, strFileType)
		
		p arrExtFiles[0].push(["a"])
		# ファイル情報 抽出
		for fixFileCnt in 0 .. (arrExtFiles.length - 1)
			puts File::stat(arrExtFiles[fixFileCnt]).size
			puts File::stat(arrExtFiles[fixFileCnt]).size.class
			puts File::stat(arrExtFiles[fixFileCnt]).size.to_s.class
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).size.to_s)		# ファイルサイズ
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).atime.to_s)	# 最終アクセス時刻
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).mtime.to_s)	# 最終更新時刻
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).ftype.to_s)	# ファイルタイプ
		end
	end
