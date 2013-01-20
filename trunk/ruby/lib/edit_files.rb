#! /usr/bin/env ruby
# =================================================
#	$Brief	�t�@�C���̒ǉ��E�폜���s�� $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/edit_files.rb $
#	
# =================================================

require "./util.rb"

	# ===============================================================
	# @brief	�w��p�X�z���̃t�H���_���폜����
	#
	# @param	strTargetPath	[in]	String	���̓t�@�C���p�X
	# 
	# @retval	�Ȃ�
	# 
	# @note		TODO �F�����I
	# ===============================================================
	#def del_directry(strTargetPath)
	#	
	#	# �p�����[�^�`�F�b�N
	#	check_param(strTargetPath)
	#	strTargetPath.gsub!("\\","/")
	#	
	#	# �{����
	#	if File.exists?(strTargetPath)
	#		# �T�u�f�B���N�g�����K�w���[�����Ƀ\�[�g�����z����쐬
	#		arrDirList = Dir::glob(strTargetPath + "**/").sort {
	#			|a,b| b.split('/').size <=> a.split('/').size
	#		}
	#		
	#		# �T�u�f�B���N�g���z���̑S�t�@�C�����폜��A�T�u�f�B���N�g�����폜
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
	# @brief	�w��p�X�z���Ƀt�H���_���쐬����
	#
	# @param	strTargetPath	[in]	String	���̓t�@�C���p�X
	# @param	strDirName		[in]	String	�쐬����f�B���N�g����
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
	# ===============================================================
	def cre_directry(strTargetPath, strDirName)
		
		# �p�����[�^�`�F�b�N
		check_param(strTargetPath, strDirName)
		strTargetPath.gsub!("\\","/")
		strDirName.gsub!("\\","/")
		arrDirName = strDirName.split("/")
		
		# �{����
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
