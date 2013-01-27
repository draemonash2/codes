#! /usr/bin/env ruby
# =================================================
#   $Brief  ���y�t�@�C���ҏW�����{���� $
#   
#   $Date:: 2013-01-11 21:00:09 +0900#$
#   $Rev: 29 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/edit_mp3.rb $
#   
#   $UsageRule:
#       ruby exe_grepReplaceMultiLine.rb <replace_target_dir_path> <replace_line_num> <search_word> <replace_file_path>
#           <replace_target_dir_path>   : �u���Ώۃf�B���N�g���p�X
#           <replace_line_num>          : �u���s��
#           <search_word>               : ����������
#           <replace_file_path>         : �u���s�t�@�C���p�X
#           <replace_file_type>         : �u���Ώۃt�@�C�����
#   
#   $Note:
#       �Ȃ�
#   
# =================================================

# =================================================
# Require �w��
# =================================================
require "./lib/input_file.rb"

# =================================
# �p�����[�^�ݒ�
# =================================
EDIT_TARGET_PATH    = "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
JACKET_DIR_PATH     = "D:/200_Pictures/30_CD Jacket"

# =================================
# ���s����
# =================================
# ���O�m�F����
def check_preProcess()
    # JACKET_DIR_PATH  �����݂��邩�m�F
    # EDIT_TARGET_PATH �����݂��邩�m�F
    # EDIT_TARGET_PATH �z���̃t�H���_�� �摜�t�@�C�����i�[����Ă��邱�Ƃ��m�F
    # TODO
end

# �摜�ҏW����
def edit_downloadJackets()
    # JACKET_DIR_PATH ����摜�t�@�C���̍ŏI�t�@�C�������擾
    arrFileNames        = Array.new()
    strFileType         = ".jpg"
    fixMaxFileNumber    = 0
    extract_file_path(JACKET_DIR_PATH, arrFileNames, strFileType)
    for fixFileNameCnt in 0 .. (arrFileNames.length - 1)
        if arrFileNames[fixFileNameCnt] =~ /Folder(...)\.jpg/
            if fixMaxFileNumber < $1.to_i
                fixMaxFileNumber = $1.to_i
            end
        end
    end
    
    # EDIT_TARGET_PATH �z���Ɋi�[���ꂽ �W���P�b�g �̃p�X���擾
    arrEditJacketPath   = Array.new()
    strFileType         = ".jpg"
    extract_file_path(EDIT_TARGET_PATH, arrEditJacketPath, strFileType)
    
    # AlbumArt ���폜
    arrEditJacketPath.delete_if {|strLine| strLine =~ /AlbumArt_.*\.jpg/ }

    puts arrEditJacketPath
    
    # EDIT_TARGET_PATH �z���Ɋi�[���ꂽ �W���P�b�g �� Folder.jpg �Ƀ��l�[��
    for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
        strOldName = arrEditJacketPath[fixEditJacketPathCnt]
        arrEditJacketPath[fixEditJacketPathCnt].gsub!(File.basename(arrEditJacketPath[fixEditJacketPathCnt], ".*"), "Folder") # arrEditJacketPath �z������l�[��
        strNewName = arrEditJacketPath[fixEditJacketPathCnt]
        puts strOldName
        puts strNewName
        File.rename(strOldName, strNewName)
    end
    
#   # EDIT_TARGET_PATH �z���Ɋi�[���ꂽ Folder.jpg �� FolderXXX.jpg �Ƀ��l�[�����A�R�s�[
#   fixMaxFileNumber += 1
#   for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
#       strOldName = arrEditJacketPath[fixEditJacketPathCnt]
#       strNewName = File.dirname(arrEditJacketPath[fixEditJacketPathCnt]) + "/" + File.basename(arrEditJacketPath[fixEditJacketPathCnt], ".*").gsub("Folder", "Folder#{format("%3d", fixMaxFileNumber)}")
#       File.rename(strOldName, strNewName)
#       fixMaxFileNumber += 1
#   end
#   
#   # Folder.jpg �� �B���t�@�C����
#   for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
#       system("attrib #{arrEditJacketPath[fixEditJacketPathCnt]} +h")
#   end
end

# MP3Gain ���s
def execute_MP3Gein()
end

# =================================
# �{����
# =================================

# Amazon�ɂăW���P�b�g�E�^�O���擾
# EditMP3Files �t�H���_�ɉ��y�t�@�C�����ړ�
# �擾�����W���P�b�g���e�t�H���_�Ɋi�[
# SuperTagEditor �ɂă^�O������

# ���O�m�F����
check_preProcess()

# �摜�ҏW����
edit_downloadJackets()

# MP3Gain ���s
execute_MP3Gein()

# Mp3tag ���s

