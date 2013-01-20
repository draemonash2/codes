#! /usr/bin/env ruby
# =================================================
#	$Brief	ファイルの追加・削除を行う $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/edit_files.rb $
#	
# =================================================

require "./util.rb"

	# ===============================================================
	# @brief	指定パス配下のフォルダを削除する
	#
	# @param	strTargetPath	[in]	String	入力ファイルパス
	# 
	# @retval	なし
	# 
	# @note		TODO ：完成！
	# ===============================================================
	#def del_directry(strTargetPath)
	#	
	#	# パラメータチェック
	#	check_param(strTargetPath)
	#	strTargetPath.gsub!("\\","/")
	#	
	#	# 本処理
	#	if File.exists?(strTargetPath)
	#		# サブディレクトリを階層が深い順にソートした配列を作成
	#		arrDirList = Dir::glob(strTargetPath + "**/").sort {
	#			|a,b| b.split('/').size <=> a.split('/').size
	#		}
	#		
	#		# サブディレクトリ配下の全ファイルを削除後、サブディレクトリを削除
	#		arrDirList.each {|strDir|
	#			Dir::foreach(strDir) {|strFile|
	#				File::delete(strDir + strFile) if ! (/\.+$/ =~ strFile)
	#			}
	#			Dir::rmdir(strDir)
	#		}
	#		Dir::rmdir(strTargetPath)
	#	else
	#		puts "already exists \"#{strTargetPath}\""
	#	end
	#end


	# ===============================================================
	# @brief	指定パス配下にフォルダを作成する
	#
	# @param	strTargetPath	[in]	String	入力ファイルパス
	# @param	strDirName		[in]	String	作成するディレクトリ名
	# 
	# @retval	なし
	# 
	# @note		なし
	# ===============================================================
	def cre_directry(strTargetPath, strDirName)
		
		# パラメータチェック
		check_param(strTargetPath, strDirName)
		strTargetPath.gsub!("\\","/")
		strDirName.gsub!("\\","/")
		arrDirName = strDirName.split("/")
		
		# 本処理
		strDirPath = strTargetPath
		for i in 0 .. (arrDirName.length - 1)
			strDirPath = strDirPath + "/" + arrDirName[i]
			if File.exists?(strDirPath)
				puts "already exists \"#{strDirPath}\""
				exit
			else
				Dir::mkdir(strDirPath)
			end
		end
	end
