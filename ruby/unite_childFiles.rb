#! /usr/bin/env ruby
# =================================================
#	$Brief	�w��f�B���N�g���z���̃t�@�C�����������āA
#			�t�@�C���ɏo�͂���
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/unite_childFiles.rb $
#	
#	$UsageRule:
#		ruby unite_childFiles.rb <dir_path> <file_type>
#		  dir_path	 : �����Ώۃf�B���N�g���p�X
#		  file_type	 : �����Ώۃt�@�C�����
#	
#	$Note:
#		�E�������ʂ� ./unite_childFiles.txt �ɏo�͂���܂�
#		
# =================================================

# =================================================
# Require �w��
# =================================================
require "./lib/input_file.rb"

# =================================================
# �p�����[�^�w��
# =================================================

# =================================================
# ���s����
# =================================================
	# ===============================================================
	# @brief	�����`�F�b�N����
	#
	# @param	�Ȃ�
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	�t�@�C�����e�𒊏o����
	#
	# @param	arrExtFileInfo		[in]	Array->Array->String	���̓t�@�C�����
	# @param	arrExtFileContents	[out]	Array->Array->String	�t�@�C�����e�i�[��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	�t�@�C������������
	#
	# @param	arrExtFileInfo		[in]	Array->Array->String	���̓t�@�C�����
	# @param	arrExtFileContents	[in]	Array->Array->String	���̓t�@�C�����e
	# @param	arrUniteFile		[out]	Array->String			�������ʊi�[��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
	# ===============================================================
	def unite_files(arrExtFileInfo, arrExtFileContents, arrUniteFile)
		# �p�����[�^�`�F�b�N
		if arrExtFileInfo.length != arrExtFileContents.length
			puts "parameter length error!"
			puts __FILE__
			puts __LINE__
			STDIN.gets()
			exit
		end
		
		for fixFileCnt in 0 .. (arrExtFileInfo.length - 1)
			# �t�@�C�����
			arrUniteFile.push(format("file_size    = %30s", arrExtFileInfo[fixFileCnt][1])) # �t�@�C���T�C�Y
			arrUniteFile.push(format("acc_time     = %30s", arrExtFileInfo[fixFileCnt][2])) # �ŏI�A�N�Z�X����
			arrUniteFile.push(format("update_time  = %30s", arrExtFileInfo[fixFileCnt][3])) # �ŏI�X�V����
			arrUniteFile.push(format("file_type    = %30s", arrExtFileInfo[fixFileCnt][4])) # �t�@�C���^�C�v
			# �t�@�C�����e
			for fixLineCnt in 0 .. (arrExtFileContents[fixFileCnt].length - 1)
				arrUniteFile.push(arrExtFileContents[fixFileCnt][fixLineCnt])
			end
		end
	end

# =================================================
# �{����
# =================================================

strDirPath	= ARGV[0].gsub("\\", "/")
strFileType = ARGV[1]

arrExtFileInfo		= Array.new() # �t�@�C�����
arrExtFileContents	= Array.new() # �t�@�C�����e
arrUniteFile		= Array.new() # ��������

strTgtPath	= File.basename(__FILE__) + ".txt"

# �����`�F�b�N
check_type()

# �t�@�C����� (�t�@�C���p�X �t�@�C���T�C�Y �ŏI�A�N�Z�X�� �ŏI�X�V���� �t�@�C���^�C�v) ���o
extract_file_info(strDirPath, arrExtFileInfo, strFileType)

# �t�@�C�����e���o
extract_file_contents(arrExtFileInfo, arrExtFileContents)

# �t�@�C������
unite_files(arrExtFileInfo, arrExtFileContents, arrUniteFile)

# �t�@�C���o��
output_txt(arrOutputArr, strTgtPath)
