#! /usr/bin/env ruby
# =================================================
#	$Brief	�z��̑�����s�� $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author:  $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/arr.rb $
#	
# =================================================

	# ===============================================================
	# @brief	�񎟌��z��̍s�Ɨ�����ւ���
	#
	# @param	arrInputArray	[in,out]	Array->Array->XXXXX	���̓f�[�^�z��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
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
	# @brief	�񎟌��z��̗v�f�S�� ���l�^�ɕϊ�����
	#
	# @param	arrInputArray	[in,out]	Array->Array->String	���̓f�[�^�z��
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
	# ===============================================================
	def conv_array_to_i!(arrInputArray)
		for i in 0 .. (arrInputArray.length - 1)
			for j in 0 .. (arrInputArray[0].length - 1)
				arrInputArray[i][j] = arrInputArray[i][j].to_i
			end
		end
	end

	# ===============================================================
	# @brief	�z��̃T�C�Y��ԋp
	#
	# @param	arrInputArray	[in]	Array->Array->XXXXX	���̓f�[�^�z��
	# 
	# @retval	arrInputArray.length
	# @retval	arrInputArray[0].length
	# 
	# @note		�Ȃ�
	# ===============================================================
	def mes_array(arrInputArray)
		return arrInputArray.length, arrInputArray[0].length
	end
	
	# ===============================================================
	# @brief	�ꎟ���z���񎟌��z��֕ϊ�
	#
	# @param	arrInputArray	[in,out]	Array->XXXXX	���̓f�[�^�z��
	#									��	Array->Array->XXXXX
	# 
	# @retval	�Ȃ�
	# 
	# @note		�Ȃ�
	# ===============================================================
	def conv_multi_array(arrInputArray)
		for fixArrayCnt in 0 .. (arrInputArray.length - 1)
		end
	end
	
	# ===============================================================
	# @brief	�ꎟ���z�����͂��A�w�肵���͈͂�u������B
	#			CSV�t�@�C�����œǂݍ���
	#			�͈͑I����͂́u����������(���K�\��)�v�Ɓu�s���v�Ƃ���
	#
	# @param	arrReplaceFor	[in,out]	Array->String	�u���f�[�^�z��
	# @param	fixReplaceLine	[in]		Fixnum			�s��
	# @param	matchSearchLine	[in]		Match			����������
	# @param	arrReplaceBase	[in]		Array->String	�u��������������
	# 
	# @retval	�Ȃ�
	# 
	# @note		�EfixReplaceLine �� 0 �ȏ���w�肷�邱��
	#			TODO : ��O�̋L�@���o����
	# ===============================================================
	def replace_lines_byWord(arrReplaceFor, fixReplaceLine, matchSearchLine, arrReplaceBase)
		
		# �p�����[�^�`�F�b�N
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
		
		# �s�u��
		for fixLineCnt in 0 .. (arrReplaceFor.length - 1)
			if arrReplaceFor[fixLineCnt].match(matchSearchLine) # ����������Ƀ}�b�`�����ꍇ
				# �s�����폜
				for fixDeleteCnt in 0 .. (fixReplaceLine - 1)
					arrReplaceFor.delete_at(fixLineCnt)
				end
				# �s�}��
				for fixInsertCnt in (0 .. (arrReplaceBase.length - 1)).reverse_each # �t��
					arrReplaceFor.insert(fixLineCnt, arrReplaceBase[fixInsertCnt])
				end
			else
				# None
			end
		end
	end
	
	# �f�o�b�O�p�֐�
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
		# ��O�����������Ƃ��̏���
	else
		# ��O���������Ȃ������Ƃ��Ɏ��s����鏈��
	ensure
		# ��O�̔����L���Ɋւ�炸�Ō�ɕK�����s���鏈��
	end
