#! /usr/bin/env ruby
# =================================================
#	$Brief	配列の操作を行う $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author:  $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/arr.rb $
#	
# =================================================

	# ===============================================================
	# @brief	二次元配列の行と列を入れ替える
	#
	# @param	arrInputArray	[in,out]	Array->Array->XXXXX	入力データ配列
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def chg_array!(arrInputArray)
		
		arrOutputArray 	= Array.new()
		
		arrOutputArray = arrInputArray.transpose
		
		arrInputArray.clear
		for i in 0 .. arrOutputArray.length - 1
			arrInputArray.push(arrOutputArray[i])
		end
	end

	# ===============================================================
	# @brief	二次元配列の要素全て 数値型に変換する
	#
	# @param	arrInputArray	[in,out]	Array->Array->String	入力データ配列
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def conv_array_to_i!(arrInputArray)
		for i in 0 .. (arrInputArray.length - 1)
			for j in 0 .. (arrInputArray[0].length - 1)
				arrInputArray[i][j] = arrInputArray[i][j].to_i
			end
		end
	end

	# ===============================================================
	# @brief	配列のサイズを返却
	#
	# @param	arrInputArray	[in]	Array->Array->XXXXX	入力データ配列
	# 
	# @retval	arrInputArray.length
	# @retval	arrInputArray[0].length
	# 
	# @note		なし
	# ===============================================================
	def mes_array(arrInputArray)
		return arrInputArray.length, arrInputArray[0].length
	end
	
	# ===============================================================
	# @brief	一次元配列を二次元配列へ変換
	#
	# @param	arrInputArray	[in,out]	Array->XXXXX	入力データ配列
	#									⇒	Array->Array->XXXXX
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def conv_multi_array(arrInputArray)
		for fixArrayCnt in 0 .. (arrInputArray.length - 1)
		end
	end
	
	# ===============================================================
	# @brief	一次元配列を入力し、指定した範囲を置換する。
	#			CSVファイル等で読み込んだ
	#			範囲選択入力は「検索文字列(正規表現)」と「行数」とする
	#
	# @param	arrReplaceFor	[in,out]	Array->String	置換データ配列
	# @param	fixReplaceLine	[in]		Fixnum			行数
	# @param	matchSearchLine	[in]		Match			検索文字列
	# @param	arrReplaceBase	[in]		Array->String	置換したい文字列
	# 
	# @retval	なし
	# 
	# @note		・fixReplaceLine は 0 以上を指定すること
	#			TODO : 例外の記法を覚える
	# ===============================================================
	def replace_lines_byWord(arrReplaceFor, fixReplaceLine, matchSearchLine, arrReplaceBase)
		
		# パラメータチェック
		if fixReplaceLine <= 0
			raise "param error!"
		#	puts "error! at #{__method__}"
		#	STDIN.gets
		#	exit
		end
		
	#	if fixReplaceLine <= 0
	#		puts "error! at #{__method__}"
	#		STDIN.gets
	#		exit
	#	end
		
		# 行置換
		for fixLineCnt in 0 .. (arrReplaceFor.length - 1)
			if arrReplaceFor[fixLineCnt].match(matchSearchLine) # 検索文字列にマッチした場合
				# 行数分削除
				for fixDeleteCnt in 0 .. (fixReplaceLine - 1)
					arrReplaceFor.delete_at(fixLineCnt)
				end
				# 行挿入
				for fixInsertCnt in (0 .. (arrReplaceBase.length - 1)).reverse_each # 逆順
					arrReplaceFor.insert(fixLineCnt, arrReplaceBase[fixInsertCnt])
				end
			else
				# None
			end
		end
	end
	
	# デバッグ用関数
	def test
		arrReplaceFor	=	[	
								"1",
								"2",
								"3",
								"4",
								"5",
								"6",
								"7",
								"8"
							]
		fixReplaceLine	=	0
		matchSearchLine =	/7/
		arrReplaceBase	=	[	
								"a",
								"b",
								"c",
								"d",
								"e"
							]
		replace_lines(arrReplaceFor, fixReplaceLine, matchSearchLine, arrReplaceBase)
	end
	
	begin
		test()
	rescue
		puts "error"
		# 例外が発生したときの処理
	else
		# 例外が発生しなかったときに実行される処理
	ensure
		# 例外の発生有無に関わらず最後に必ず実行する処理
	end
