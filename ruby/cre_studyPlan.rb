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
#       ruby cre_studyPlan.rb <sep_page_num> <input_csv_path>
#         sep_page_num   : 分割ページ数
#         input_csv_path : 入力する CSV ファイルパス
#   
#   $Note:
#       TODO: 行列反転処理を削除する。
#       TODO: 入力する CSV に sep_page_num 記述対応
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
    def join_inputArray(arrJoinedArray, arrChgedArray)
        # 結合
        for i in 0 .. (arrChgedArray.length - 1)
            arrJoinedArray.concat(arrChgedArray[i])
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
    def cre_separated_plan(arrOutputPlan, arrJoinedArray, fixSepPageNum)
        
        # ページ数算出
        fixStartPageNum = arrJoinedArray[0]
        fixLastPageNum  = arrJoinedArray[arrJoinedArray.length - 1]
        
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
    # @param    arrOutputArray   [out]   array->array->integer    分割結果格納領域
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
    fixSepPageNum   = ARGV[0].to_i
    strInpPath      = ARGV[1].gsub("\\", "/")
    strOutPath      = strInpPath.gsub("input", "output")
    arrInputArray   = Array.new()
    arrJoinedArray  = Array.new()       # 結合後出力
    arrOutputArray  = Array.new()       # 最終出力結果
    
    # CSV 入力
    input_csv(strInpPath, arrInputArray)
    
    # 入力ファイルチェック
    chk_inputcsv(arrInputArray)
    
    # 説明行削除
    arrInputArray.delete_at(0)
    
    # 行列反転
    chg_array!(arrInputArray)
    
    # 全て数値に変換
    conv_array_to_i!(arrInputArray)
    
    # 全ての配列を結合
    join_inputArray(arrJoinedArray, arrInputArray)
    
    # 結合した配列情報をもとに、単位ページで分割
    # 計算結果を配列で返す
    cre_separated_plan(arrOutputArray, arrJoinedArray, fixSepPageNum)
    
    # ページ数、ページ範囲を挿入
    add_pageInfo!(arrOutputArray)
    
    # 行列反転
    chg_array!(arrOutputArray)
    
    # 配列を出力
    output_csv(strOutPath, arrOutputArray, "w")
    
