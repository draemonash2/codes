#! /usr/bin/env ruby
# =================================================
#	$Brief	�t�@�C���̓��͂��s�� $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/input_file.rb $
#	
# =================================================

require "find"

	# ===============================================================
	# @brief	CSV�t�@�C������͂��A�񎟌��z��Ɋi�[����
	#
	# @param	strTgtPath		[in]	String					���̓t�@�C���p�X
	# @param	arrInputArray	[in]	Array->Array->String	���̓f�[�^�z��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	TSV�t�@�C������͂��A�񎟌��z��Ɋi�[����
	#
	# @param	strTgtPath		[in]	String					���̓t�@�C���p�X
	# @param	arrInputArray	[in]	Array->Array->String	���̓f�[�^�z��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	TXT �t�@�C������͂��A�z��Ɋi�[����
	#
	# @param	strTargetPath	[in]	String			���̓t�@�C���p�X
	# @param	arrInputArray	[out]	Array->String	���̓f�[�^�z��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	�t�@�C���𒊏o����
	#			�f�B���N�g���͒��o���Ȃ�
	#
	# @param	strTargetPath	[in]	String			���̓t�@�C���p�X
	# @param	arrExtFiles		[out]	Array->String	���̓f�[�^�z��
	# @param	strFileType		[in]	String			���̓t�@�C���^�C�v
	#				none	: �t�@�C����ʎw��Ȃ�
	#				.xx		: �t�@�C����ʎw�肠��( xx => �g���q��)
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	�t�@�C���𒊏o����
	#			�f�B���N�g���͒��o���Ȃ�
	#			�����Ƀt�@�C�����𒊏o����
	#				�E�t�@�C���p�X
	#				�E�t�@�C���T�C�Y
	#				�E�ŏI�A�N�Z�X����
	#				�E�ŏI�X�V����
	#				�E�t�@�C���^�C�v
	#
	# @param	strTargetPath	[in]	String					���̓t�@�C���p�X
	# @param	arrExtFiles		[out]	Array->Array->String	���̓f�[�^�z��
	# @param	strFileType		[in]	String					���̓t�@�C���^�C�v
	#				none	: �t�@�C����ʎw��Ȃ�
	#				.xx		: �t�@�C����ʎw�肠��( xx => �g���q��)
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
	# ===============================================================
	def extract_file_info(strTargetPath, arrExtFiles, strFileType)
		
		# �t�@�C���p�X ���o
		extract_file_path(strTargetPath, arrExtFiles, strFileType)
		
		p arrExtFiles[0].push(["a"])
		# �t�@�C����� ���o
		for fixFileCnt in 0 .. (arrExtFiles.length - 1)
			puts File::stat(arrExtFiles[fixFileCnt]).size
			puts File::stat(arrExtFiles[fixFileCnt]).size.class
			puts File::stat(arrExtFiles[fixFileCnt]).size.to_s.class
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).size.to_s)		# �t�@�C���T�C�Y
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).atime.to_s)	# �ŏI�A�N�Z�X����
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).mtime.to_s)	# �ŏI�X�V����
			arrExtFiles[fixFileCnt].push(File::stat(arrExtFiles[fixFileCnt]).ftype.to_s)	# �t�@�C���^�C�v
		end
	end
