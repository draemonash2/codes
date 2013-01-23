#! /usr/bin/env ruby
# =================================================
#	$Brief	�V���b�g�_�E�����Ɏ��s���鏈�����L�q $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author:  $
#	$HeadURL: file:///C:/Repo/trunk/ruby/dig_ToDoAtShutDown.rb $
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
require "find"

# =================================
# �p�����[�^�ݒ�
# =================================
DESKTOP_PATH	= "#{ENV['USERPROFILE']}/Desktop".gsub("\\","/")

# =================================
# ���s����
# =================================
# RightsNetworkMediaPlugIn ���폜����
def delete_RightsNetworkMediaPlugIn()
	Find.find(DESKTOP_PATH) {|strFilePath|
		if strFilePath =~ /RightsNetworkMediaPlugIn.*\.exe/
			File.delete(strFilePath)
		end
	}
end

# �R�~�b�g & �v�b�V��
def push_automaticGithub(strDirPath, strPushDir)
#	git add # �ݒ�t�@�C���̓t�@�C���������邱�Ƃ��Ȃ����� add ���Ȃ�
#	TODO �J�����g�f�B���N�g���ύX�����ǉ�
	system("git commit -m \"�����R�~�b�g\"")
	system("git push #{strPushDir}")
end

# =================================
# �{����
# =================================
# RightsNetworkMediaPlugIn ���폜����
delete_RightsNetworkMediaPlugIn()

# Vim Setting �R�~�b�g & �v�b�V��
push_automaticGithub(strDirPath, strPushDir)

# X-Finder Setting �R�~�b�g & �v�b�V��
push_automaticGithub(strDirPath, strPushDir)
