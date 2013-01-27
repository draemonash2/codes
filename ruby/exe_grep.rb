#! /usr/bin/env ruby
# =================================================
#   $Brief  連続Grepを可能にする $
#   
#   $Date:: 2013-01-07 00:30:23 +0900#$
#   $Rev: 28 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/exe_grep.rb $
#   
#   $UsageRule:
#   
#   $Note:
#       なし
#       
# =================================================

# =================================================
# Require 指定
# =================================================
require "find"

# =================================================
# パラメータ指定
# =================================================

# =================================================
# 実行処理
# =================================================
def check_param_rb(fixArgvLen)
    if ARGV.length != fixArgvLen
        puts "few argv len error!"
    end
end

# =====================================
# 本処理
# =====================================

# 変数定義
strRootDirPath      = ARGV[0].gsub("\\","/")
strSearchWord       = ARGV[1]
fixMatchNum         = 0
arrMatchLine        = []

# パラメータチェック
check_param_rb(2)

arrMatchLine.push("search directry is ...")
arrMatchLine.push("           #{strRootDirPath}")
arrMatchLine.push("search word is \"#{strSearchWord}\"")


# Grep 処理
Find.find(strRootDirPath) { |strFilePath|
    if File.ftype(strFilePath) == "file"
        fileFilePath = $stdin
        fileFilePath = File.open(strFilePath, "r")
        arrGetLine = fileFilePath.readlines
        for fixLine in 0 .. (arrGetLine.length - 1)
            if arrGetLine[fixLine] =~ Regexp.new(strSearchWord)
                strPushLine = format("   %-50s [%05d]", arrGetLine[fixLine].chomp, fixLine + 1)
                arrMatchLine.push(strPushLine)
                fixMatchNum += 1
            end
        end
        fileFilePath.close
    end
}
puts arrMatchLine
puts "match num #{fixMatchNum}"

