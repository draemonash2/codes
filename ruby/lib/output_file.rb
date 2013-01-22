#! /usr/bin/env ruby
# =================================================
#	$Brief	ファイルの出力を行う $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/output_file.rb $
#	
# =================================================

	# ===============================================================
	# @brief	指定された二次元配列を、CSVファイルに出力する
	#
	# @param	arrOutputArr	[in]	Array->Array->String	入力データ配列
	# @param	strTgtPath		[in]	String					入力ファイルパス
	# 
	# @retval	なし
	# 
	# @note		・既に strTgtPath ファイルが存在する場合は、
	# 			  ファイル名に "_XXX" を付与して出力
	#			・strTgtPath の拡張子は確認しない
	# ===============================================================
	def output_csv(arrOutputArr, strTgtPath)
		# ファイルパス 書き換え
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i].join(",")
		end
	end
	
	# ===============================================================
	# @brief	指定された二次元配列を、TSVファイル形式として出力する
	#
	# @param	arrOutputArr	[in]	Array->Array->String	入力データ配列
	# @param	strTgtPath		[in]	String					入力ファイルパス
	# 
	# @retval	なし
	# 
	# @note		・既に strTgtPath ファイルが存在する場合は、
	# 			  ファイル名に "_XXX" を付与して出力
	#			・strTgtPath の拡張子は確認しない
	# ===============================================================
	def output_tsv(arrOutputArr, strTgtPath)
		# ファイルパス 書き換え
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i].join("\t")
		end
	end
	
	# ===============================================================
	# @brief	指定された配列を、TXT ファイルに出力する
	#
	# @param	arrOutputArr	[in]	Array->String	入力データ配列
	# @param	strTgtPath		[in]	String			入力ファイルパス
	# 
	# @retval	なし
	# 
	# @note		・既に strTgtPath ファイルが存在する場合は、
	# 			  ファイル名に "_XXX" を付与して出力
	#			・strTgtPath の拡張子は確認しない
	# ===============================================================
	def output_txt(arrOutputArr, strTgtPath)
		# ファイルパス 書き換え
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i]
		end
	end
	
	# ===============================================================
	# @brief	ファイルパスを確認し、ファイルを作成する
	#			もし同ファイル名が存在する場合、「_XXX」を付与して作成する
	#
	# @param	strTgtPath		[in]	String	入力ファイルパス
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def convert_output_file_name(strTgtPath)
		fixFileNum = 1
		strTgtPath.gsub!("\\","/")
		while File.exists?(strTgtPath)
			arrFileName = Array.new()
			arrFileName = strTgtPath.split(".")
			strTgtPath.gsub!(/.*/, (arrFileName[0] + "_" + format("%03d", fixFileNum) + "." + arrFileName[1]))
			fixFileNum += 1
		end
