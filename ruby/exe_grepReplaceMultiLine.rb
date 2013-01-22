#! /usr/bin/env ruby
# =================================================
#	$Brief	�����s�̒u�������{���� $
#	
#	$Date:: 2013-01-11 21:00:09 +0900#$
#	$Rev: 29 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/exe_grepReplaceMultiLine.rb $
#	
#	$UsageRule:
#		ruby exe_grepReplaceMultiLine.rb <replace_target_dir_path> <replace_line_num> <search_word> <replace_file_path>
#			<replace_target_dir_path>	: �u���Ώۃf�B���N�g���p�X
#			<replace_line_num>			: �u���s��
#			<search_word>				: ����������
#			<replace_file_path>			: �u���s�t�@�C���p�X
#			<replace_file_type>			: �u���Ώۃt�@�C�����
#	
#	$Note:
#		�Ȃ�
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

# =================================================
# �{����
# =================================================
strReplaceTargetDirPath	= ARGV[0].gsub("\\","/")
strReplaceFilePath		= ARGV[1].gsub("\\","/")
fixReplaceLine			= ARGV[2].to_i
matchSearchLine			= Regexp.new(ARGV[3])
strFileType				= ARGV[4]

# �u���Ώۃt�@�C���ꗗ���o
arrExtFiles	= Array.new()
if strFileType == "none"
	extract_file_path(strReplaceTargetDirPath, arrExtFiles, "none")
else
	extract_file_path(strReplaceTargetDirPath, arrExtFiles, strFileType)
end

# �u���s�����
arrReplaceTxtFile	= Array.new()
input_txt(strReplaceFilePath, arrReplaceTxtFile)

# �^�[�Q�b�g�t�H���_�z���̈ꊇ�u��
for fixFileCnt in 0 .. (arrExtFiles.length - 1)
	# �u���Ώۃt�@�C�������
	arrReplaceTargetFile	= Array.new()
	input_txt(strReplaceTargetDirPath, arrReplaceTargetFile)
	
	# �u��
	replace_lines_byWord(arrReplaceTargetFile, fixReplaceLine, matchSearchLine, arrReplaceTxtFile)
end
