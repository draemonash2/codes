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
#		ruby dig_ToDoAtShutDown.rb
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
	Dir::chdir(strDirPath)
	system("git add -u \.")					# �����Ώۃt�@�C���̂݃R�~�b�g�ΏۂƂ���
	system("git commit -m \"Auto Commit\"")	# �ύX�̃R�~�b�g
	system("git push #{strPushDir}")
end

# =================================
# �{����
# =================================
	# RightsNetworkMediaPlugIn ���폜����
	delete_RightsNetworkMediaPlugIn()

	# Setting �R�~�b�g & �v�b�V��
	strDirPath	= "C:/prg"
	strPushDir	= "setting"
	push_automaticGithub(strDirPath, strPushDir)
