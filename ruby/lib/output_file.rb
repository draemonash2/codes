#! /usr/bin/env ruby
# =================================================
#	$Brief	�t�@�C���̏o�͂��s�� $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/output_file.rb $
#	
# =================================================

	# ===============================================================
	# @brief	�w�肳�ꂽ�񎟌��z����ACSV�t�@�C���ɏo�͂���
	#
	# @param	arrOutputArr	[in]	Array->Array->String	���̓f�[�^�z��
	# @param	strTgtPath		[in]	String					���̓t�@�C���p�X
	# 
	# @retval	�Ȃ�
	# 
	# @note		�E���� strTgtPath �t�@�C�������݂���ꍇ�́A
	# 			  �t�@�C������ "_XXX" ��t�^���ďo��
	#			�EstrTgtPath �̊g���q�͊m�F���Ȃ�
	# ===============================================================
	def output_csv(arrOutputArr, strTgtPath)
		# �t�@�C���p�X ��������
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i].join(",")
		end
	end
	
	# ===============================================================
	# @brief	�w�肳�ꂽ�񎟌��z����ATSV�t�@�C���`���Ƃ��ďo�͂���
	#
	# @param	arrOutputArr	[in]	Array->Array->String	���̓f�[�^�z��
	# @param	strTgtPath		[in]	String					���̓t�@�C���p�X
	# 
	# @retval	�Ȃ�
	# 
	# @note		�E���� strTgtPath �t�@�C�������݂���ꍇ�́A
	# 			  �t�@�C������ "_XXX" ��t�^���ďo��
	#			�EstrTgtPath �̊g���q�͊m�F���Ȃ�
	# ===============================================================
	def output_tsv(arrOutputArr, strTgtPath)
		# �t�@�C���p�X ��������
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i].join("\t")
		end
	end
	
	# ===============================================================
	# @brief	�w�肳�ꂽ�z����ATXT �t�@�C���ɏo�͂���
	#
	# @param	arrOutputArr	[in]	Array->String	���̓f�[�^�z��
	# @param	strTgtPath		[in]	String			���̓t�@�C���p�X
	# 
	# @retval	�Ȃ�
	# 
	# @note		�E���� strTgtPath �t�@�C�������݂���ꍇ�́A
	# 			  �t�@�C������ "_XXX" ��t�^���ďo��
	#			�EstrTgtPath �̊g���q�͊m�F���Ȃ�
	# ===============================================================
	def output_txt(arrOutputArr, strTgtPath)
		# �t�@�C���p�X ��������
		convert_output_file_name(strTgtPath)
		
		output_file = $stdout
		output_file = File.open(strTgtPath, 'w')
		
		for i in 0..arrOutputArr.length - 1
			output_file.puts arrOutputArr[i]
		end
	end
	
	# ===============================================================
	# @brief	�t�@�C���p�X���m�F���A�t�@�C�����쐬����
	#			�������t�@�C���������݂���ꍇ�A�u_XXX�v��t�^���č쐬����
	#
	# @param	strTgtPath		[in]	String	���̓t�@�C���p�X
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
