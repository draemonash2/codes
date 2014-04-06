#! /usr/bin/env ruby
# =================================================
#   $Brief  ファイルの出力を行う $
#   
#   $Date:: 2013-01-11 21:00:09 +0900#$
#   $Rev: 29 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/lib/output_file.rb $
#   
# =================================================

    # ===============================================================
    # @brief    指定された二次元配列を、CSVファイルに出力する
    #
    # @param    strTargetPath   [in]    String                  入力ファイルパス
    # @param    arrOutputArr    [in]    Array->Array->String    入力データ配列
    # @param    strWriteMode    [in]    String  書き込みモード
    #               w : 新規作成書込みモード
    #                   既存ファイルを指定した場合、ファイル名に "_XXX" を付与して出力
    #               a : 上書きモード
    #                   既存ファイルを指定した場合、上書きする。
    # 
    # @retval   なし
    # 
    # @note     ・strTargetPath の拡張子は確認しない
    # ===============================================================
    def output_csv(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # ファイルパス 書き換え
            when "a" then   # None
            else            raise "write mode error!"
        end
        
        output_file = $stdout
        output_file = File.open(strTargetPath, 'w')
        
        for i in 0..arrOutputArr.length - 1
            output_file.puts arrOutputArr[i].join(",")
        end
    end
    
    # ===============================================================
    # @brief    指定された二次元配列を、TSVファイル形式として出力する
    #
    # @param    strTargetPath   [in]    String                  入力ファイルパス
    # @param    arrOutputArr    [in]    Array->Array->String    入力データ配列
    # @param    strWriteMode    [in]    String  書き込みモード
    #               w : 新規作成書込みモード
    #                   既存ファイルを指定した場合、ファイル名に "_XXX" を付与して出力
    #               a : 上書きモード
    #                   既存ファイルを指定した場合、上書きする。
    # 
    # @retval   なし
    # 
    # @note     ・strTargetPath の拡張子は確認しない
    # ===============================================================
    def output_tsv(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # ファイルパス 書き換え
            when "a" then   # None
            else            raise "write mode error!"
        end
        
        output_file = $stdout
        output_file = File.open(strTargetPath, 'w')
        
        for i in 0..arrOutputArr.length - 1
            output_file.puts arrOutputArr[i].join("\t")
        end
    end
    
    # ===============================================================
    # @brief    指定された配列を、TXT ファイルに出力する
    #
    # @param    strTargetPath   [in]    String          入力ファイルパス
    # @param    arrOutputArr    [in]    Array->String   入力データ配列
    # @param    strWriteMode    [in]    String          書き込みモード
    #               w : 新規作成書込みモード
    #                   既存ファイルを指定した場合、ファイル名に "_XXX" を付与して出力
    #               a : 上書きモード
    #                   既存ファイルを指定した場合、上書きする。
    # 
    # @retval   なし
    # 
    # @note     ・strTargetPath の拡張子は確認しない
    # ===============================================================
    def output_txt(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # ファイルパス 書き換え
            when "a" then   # None
            else            raise "write mode error!"
        end
        
        output_file = $stdout
        output_file = File.open(strTargetPath, 'w')
        
        for i in 0..arrOutputArr.length - 1
            output_file.puts arrOutputArr[i]
        end
    end
    
    # ===============================================================
    # @brief    ファイルパスを確認し、ファイルを作成する
    #           もし同ファイル名が存在する場合、「_XXX」を付与して作成する
    #
    # @param    strTargetPath   [in]    String  入力ファイルパス
    # @param    strWriteMode    [in]    String  書き込みモード
    #               w : 新規作成書込みモード
    #                   既存ファイルを指定した場合、ファイル名に "_XXX" を付与して出力
    #               a : 上書きモード
    #                   既存ファイルを指定した場合、上書きする。
    # 
    # @retval   なし
    # 
    # @note     なし
    # ===============================================================
    def convert_output_file_name(strTargetPath, strWriteMode)
        fixFileNum = 1
        strTargetPath.gsub!("\\","/")
        
        while File.exists?(strTargetPath)
            if strTargetPath =~ /_(\d\d\d)\./
                strSrc  =   $1                          + File.extname(strTargetPath)
                strDst  =   format("%03d", $1.to_i + 1) + File.extname(strTargetPath)
            else
                strSrc  =             File.extname(strTargetPath)
                strDst  =   "_001"  + File.extname(strTargetPath)
            end
            strTargetPath.gsub!(strSrc, strDst)
            fixFileNum += 1
        end
    end
    
