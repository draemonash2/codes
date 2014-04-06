#! /usr/bin/env ruby
# =============================================
#   $Brief  勉強の計画表を生成する $
#   
#   $Date:: 2013-01-07 00:30:23 +0900#$
#   $Rev: 28 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/cre_studyPlan.rb $
#   
#   $UsageRule:
#       ・実行方法
#         ruby cre_studyPlan.rb <sep_page_num> <input_csv_path>
#           sep_page_num   : 分割ページ数
#           input_csv_path : 入力する CSV ファイルパス
#         
#       ・入力ファイル形式
#         | SepPageNum | StartEndPage | Level1 | Level2 | Level3 |
#         |     10     |      12      |   43   |   12   |   12   |
#         |            |     435      |   65   |   21   |   15   |
#         |            |              |   83   |   30   |   20   |
#         |            |              |   ・   |   ・   |   ・   |
#         |            |              |   ・   |   ・   |   ・   |
#         |            |              |   ・   |   ・   |   ・   |
#         
#       ・出力ファイル形式
#         | StartPage  |   12   |   16   |   ・   |   ・   |   ・   |
#         | EndPage    |   15   |   20   |   ・   |   ・   |   ・   |
#         | PageNum    |    3   |    4   |   ・   |   ・   |   ・   |
#   
#   $Note:
#       TODO: 事前準備の関数化
#       TODO: 構造体化
#       TODO: sep_page_num 付近で分割されるように… (Level 3 を入力させればまだいいけど…)
# =============================================

# =================================================
# Require 指定
# =================================================
require "./lib/debug.rb"
require "./lib/input_file.rb"
require "./lib/output_file.rb"
require "./lib/arr.rb"

# =================================================
# 実行処理
# =================================================
    # ===============================================================
    # @brief    入力ファイルが正しくソートされているか確認する
    #           ソートされていない場合は、異常終了する
    #
    # @param    arrInputArray   [in]    array->array->string    入力配列
    # 
    # @retval   なし
    # 
    # @note     arrInputArray は行列反転、数値変換は未実施であること
    # ===============================================================
    def chk_inputcsv(arrInputArray)
        arrTemp = Array.new()
        arrTemp = arrInputArray
        
        arrTemp.delete_at(0)
        chg_array!(arrTemp)
        conv_array_to_i!(arrTemp)
        
        for i in 0 .. (arrTemp.length - 1)
            for j in 0 .. (arrTemp[i].length - 2)
                # 終了確認
                if arrTemp[i][j + 1] == 0
                    break
                end
                
                # ソート済み確認
                if arrTemp[i][j] < arrTemp[i][j + 1]
                    # None (正常)
                else
                    puts "input csv file error!"
                    puts "input file is not sorted!"
                    puts "row   = #{i}"
                    puts "line  = #{j + 2}_#{j + 3}"
                    exit
                end
            end
        end
    end
    
    def execute_pre_process()
    end

    # =====================================================
    # @brief    全ての配列を結合する
    # 
    # @param    arrJoinedArray  [out]   array->string           結合後の配列
    # @param    arrChgedArray   [in]    array->array->string    入力配列
    #
    # @retval   なし
    # 
    # @note     なし
    # =====================================================
    def join_inputArray(arrJoinedArray, arrInputArray, fixSepPageNum, fixStartPageNum, fixLastPageNum )
        # 結合
        for i in 0 .. (arrInputArray.length - 1)
            arrJoinedArray.concat(arrInputArray[i])
            arrJoinedArray.push(fixSepPageNum)
            arrJoinedArray.push(fixStartPageNum)
            arrJoinedArray.push(fixLastPageNum)
        end
        
        # 重複を排除
        arrJoinedArray.uniq!
        
        # ソート
        arrJoinedArray.sort!
        
        # 0 ページを削除
        if arrJoinedArray[0] == 0
            arrJoinedArray.delete_at(0)
        end
    end

    # =====================================================
    # @brief    結合した配列情報をもとに、単位ページで分割
    #           計算結果を配列で返す
    #       
    # @param    arrOutputPlan   [out]   array->array->string    分割結果格納領域
    # @param    arrJoinedArray  [in]    array->array->string    結合後配列
    # @param    fixSepPageNum   [in]    fixnum                  分割ページ数
    #
    # @retval   なし
    # 
    # @note     なし
    # =====================================================
    def cre_separated_plan(arrOutputPlan, arrJoinedArray, fixSepPageNum, fixStartPageNum, fixLastPageNum)
        
        fixArrIdx = 0
        fixCurPageNum = fixStartPageNum
        
        # 最初のページ出力
        arrOutputPlan.push(fixCurPageNum)   # 出力
        
        # 中間のページ出力
        while fixArrIdx < arrJoinedArray.length - 1
            
            if fixCurPageNum < arrJoinedArray[fixArrIdx]
                fixCurPageNum   += fixSepPageNum
            else
                if fixCurPageNum < arrJoinedArray[fixArrIdx + 1]
                    # 近い方
                    first   = (fixCurPageNum - arrJoinedArray[fixArrIdx]).abs
                    second  = (fixCurPageNum - arrJoinedArray[fixArrIdx + 1]).abs
                    
                    if second > fixSepPageNum
                        arrOutputPlan.push(fixCurPageNum)   # 出力
                        fixCurPageNum   += fixSepPageNum
                    else
                        if first < second
                            arrOutputPlan.push(arrJoinedArray[fixArrIdx])  # 出力
                            fixCurPageNum   = arrJoinedArray[fixArrIdx] + fixSepPageNum
                        else
                            arrOutputPlan.push(arrJoinedArray[fixArrIdx + 1])  # 出力
                            fixCurPageNum   = arrJoinedArray[fixArrIdx + 1] + fixSepPageNum
                        end
                    end
                else
                    fixArrIdx   += 1
                end
            end
        end
        
        # 最終ページ出力
        arrOutputPlan.push(fixLastPageNum)  # 出力
    end

    # =====================================================
    # @brief    ページ数とページ範囲を付加する
    # 
    # @param    arrOutputArray   [in,out]   array->array->integer    分割結果格納領域
    #
    # @retval   なし
    # 
    # @note     なし
    # =====================================================
    def add_pageInfo!(arrOutputArray)
        
        arrPageStart   = Array.new()
        arrPageNum     = Array.new()
        arrPageEnd     = Array.new()
        
        for i in 0 .. (arrOutputArray.length - 2)
            if i == 0
                arrPageStart.push ("StartPage")
                arrPageEnd.push   ("EndPage")
                arrPageNum.push   ("PageNum")
            else
                arrPageStart.push ("#{arrOutputArray[i]}")
                arrPageEnd.push   ("#{arrOutputArray[i + 1] - 1}")
                arrPageNum.push   ("#{arrOutputArray[i + 1] - arrOutputArray[i]}")
            end
        end
        
        arrOutputArray.clear
        arrOutputArray.push(arrPageStart)
        arrOutputArray.push(arrPageEnd)
        arrOutputArray.push(arrPageNum)
    end

# =================================================
# 本処理
# =================================================
    
    # 入力
    strInpPath      = ARGV[0].gsub("\\", "/")
    strOutPath      = strInpPath.gsub("input", "output")
    arrInputArray   = Array.new()
    arrJoinedArray  = Array.new()       # 結合後出力
    arrOutputArray  = Array.new()       # 最終出力結果
    
    # CSV 入力
    input_csv(strInpPath, arrInputArray)
    
    # 入力ファイルチェック
#   chk_inputcsv(arrInputArray)
#   ⇒ TODO : 要改良！
    
    # 基本情報入力
    fixSepPageNum   = arrInputArray[1][0].to_i
    fixStartPageNum = arrInputArray[1][1].to_i
    fixLastPageNum  = arrInputArray[2][1].to_i
    
    # 説明行削除
    arrInputArray.delete_at(0)
    
    # 行列反転
    chg_array!(arrInputArray)
    
    # 分割ページ数、開始終了ページ数 列削除
    arrInputArray.delete_at(0)
    arrInputArray.delete_at(0)
    
    # 全て数値に変換
    conv_array_to_i!(arrInputArray)
    
    # 全ての配列を結合
    join_inputArray(arrJoinedArray, arrInputArray, fixSepPageNum, fixStartPageNum, fixLastPageNum)
    
    # 結合した配列情報をもとに、単位ページで分割
    # 計算結果を配列で返す
    cre_separated_plan(arrOutputArray, arrJoinedArray, fixSepPageNum, fixStartPageNum, fixLastPageNum)
    
    # ページ数、ページ範囲を挿入
    add_pageInfo!(arrOutputArray)
    
    # 配列を出力
    output_csv(strOutPath, arrOutputArray, "w")
   
