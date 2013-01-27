#! /usr/bin/env ruby
# =================================================
#   $Brief  音楽ファイル編集を実施する $
#   
#   $Date:: 2013-01-11 21:00:09 +0900#$
#   $Rev: 29 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/edit_mp3.rb $
#   
#   $UsageRule:
#       ruby exe_grepReplaceMultiLine.rb <replace_target_dir_path> <replace_line_num> <search_word> <replace_file_path>
#           <replace_target_dir_path>   : 置換対象ディレクトリパス
#           <replace_line_num>          : 置換行数
#           <search_word>               : 検索文字列
#           <replace_file_path>         : 置換行ファイルパス
#           <replace_file_type>         : 置換対象ファイル種別
#   
#   $Note:
#       なし
#   
# =================================================

# =================================================
# Require 指定
# =================================================
require "./lib/input_file.rb"

# =================================
# パラメータ設定
# =================================
EDIT_TARGET_PATH    = "#{ENV['USERPROFILE']}/Desktop/EditMP3Files".gsub("\\", "/")
JACKET_DIR_PATH     = "D:/200_Pictures/30_CD Jacket"

# =================================
# 実行処理
# =================================
# 事前確認処理
def check_preProcess()
    # JACKET_DIR_PATH  が存在するか確認
    # EDIT_TARGET_PATH が存在するか確認
    # EDIT_TARGET_PATH 配下のフォルダに 画像ファイルが格納されていることを確認
    # TODO
end

# 画像編集処理
def edit_downloadJackets()
    # JACKET_DIR_PATH から画像ファイルの最終ファイル名を取得
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
    
    # EDIT_TARGET_PATH 配下に格納された ジャケット のパスを取得
    arrEditJacketPath   = Array.new()
    strFileType         = ".jpg"
    extract_file_path(EDIT_TARGET_PATH, arrEditJacketPath, strFileType)
    
    # AlbumArt を削除
    arrEditJacketPath.delete_if {|strLine| strLine =~ /AlbumArt_.*\.jpg/ }

    puts arrEditJacketPath
    
    # EDIT_TARGET_PATH 配下に格納された ジャケット を Folder.jpg にリネーム
    for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
        strOldName = arrEditJacketPath[fixEditJacketPathCnt]
        arrEditJacketPath[fixEditJacketPathCnt].gsub!(File.basename(arrEditJacketPath[fixEditJacketPathCnt], ".*"), "Folder") # arrEditJacketPath 配列をリネーム
        strNewName = arrEditJacketPath[fixEditJacketPathCnt]
        puts strOldName
        puts strNewName
        File.rename(strOldName, strNewName)
    end
    
#   # EDIT_TARGET_PATH 配下に格納された Folder.jpg を FolderXXX.jpg にリネームし、コピー
#   fixMaxFileNumber += 1
#   for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
#       strOldName = arrEditJacketPath[fixEditJacketPathCnt]
#       strNewName = File.dirname(arrEditJacketPath[fixEditJacketPathCnt]) + "/" + File.basename(arrEditJacketPath[fixEditJacketPathCnt], ".*").gsub("Folder", "Folder#{format("%3d", fixMaxFileNumber)}")
#       File.rename(strOldName, strNewName)
#       fixMaxFileNumber += 1
#   end
#   
#   # Folder.jpg を 隠しファイル化
#   for fixEditJacketPathCnt in 0 .. (arrEditJacketPath.length - 1)
#       system("attrib #{arrEditJacketPath[fixEditJacketPathCnt]} +h")
#   end
end

# MP3Gain 実行
def execute_MP3Gein()
end

# =================================
# 本処理
# =================================

# Amazonにてジャケット・タグを取得
# EditMP3Files フォルダに音楽ファイルを移動
# 取得したジャケットを各フォルダに格納
# SuperTagEditor にてタグを入れる

# 事前確認処理
check_preProcess()

# 画像編集処理
edit_downloadJackets()

# MP3Gain 実行
execute_MP3Gein()

# Mp3tag 実行

