#! /usr/bin/env ruby
# =============================================
#   $Brief  �׋��̌v��\�𐶐����� $
#   
#   $Date:: 2013-01-07 00:30:23 +0900#$
#   $Rev: 28 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/cre_studyPlan.rb $
#   
#   $UsageRule:
#       ruby cre_studyPlan.rb <sep_page_num> <input_csv_path>
#         sep_page_num   : �����y�[�W��
#         input_csv_path : ���͂��� CSV �t�@�C���p�X
#   
#   $Note:
#       TODO: �s�񔽓]�������폜����B
#       TODO: ���͂��� CSV �� sep_page_num �L�q�Ή�
#       TODO: sep_page_num �t�߂ŕ��������悤�Ɂc (Level 3 ����͂�����΂܂��������ǁc)
# =============================================

# =================================================
# Require �w��
# =================================================
require "./lib/debug.rb"
require "./lib/input_file.rb"
require "./lib/output_file.rb"
require "./lib/arr.rb"

# =================================================
# ���s����
# =================================================
    # ===============================================================
    # @brief    ���̓t�@�C�����������\�[�g����Ă��邩�m�F����
    #           �\�[�g����Ă��Ȃ��ꍇ�́A�ُ�I������
    #
    # @param    arrInputArray   [in]    array->array->string    ���͔z��
    # 
    # @retval   �Ȃ�
    # 
    # @note     arrInputArray �͍s�񔽓]�A���l�ϊ��͖����{�ł��邱��
    # ===============================================================
    def chk_inputcsv(arrInputArray)
        arrTemp = Array.new()
        arrTemp = arrInputArray
        
        arrTemp.delete_at(0)
        chg_array!(arrTemp)
        conv_array_to_i!(arrTemp)
        
        for i in 0 .. (arrTemp.length - 1)
            for j in 0 .. (arrTemp[i].length - 2)
                # �I���m�F
                if arrTemp[i][j + 1] == 0
                    break
                end
                
                # �\�[�g�ς݊m�F
                if arrTemp[i][j] < arrTemp[i][j + 1]
                    # None (����)
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
    # @brief    �S�Ă̔z�����������
    # 
    # @param    arrJoinedArray  [out]   array->string           ������̔z��
    # @param    arrChgedArray   [in]    array->array->string    ���͔z��
    #
    # @retval   �Ȃ�
    # 
    # @note     �Ȃ�
    # =====================================================
    def join_inputArray(arrJoinedArray, arrChgedArray)
        # ����
        for i in 0 .. (arrChgedArray.length - 1)
            arrJoinedArray.concat(arrChgedArray[i])
        end
        
        # �d����r��
        arrJoinedArray.uniq!
        
        # �\�[�g
        arrJoinedArray.sort!
        
        # 0 �y�[�W���폜
        if arrJoinedArray[0] == 0
            arrJoinedArray.delete_at(0)
        end
    end

    # =====================================================
    # @brief    ���������z��������ƂɁA�P�ʃy�[�W�ŕ���
    #           �v�Z���ʂ�z��ŕԂ�
    #       
    # @param    arrOutputPlan   [out]   array->array->string    �������ʊi�[�̈�
    # @param    arrJoinedArray  [in]    array->array->string    ������z��
    # @param    fixSepPageNum   [in]    fixnum                  �����y�[�W��
    #
    # @retval   �Ȃ�
    # 
    # @note     �Ȃ�
    # =====================================================
    def cre_separated_plan(arrOutputPlan, arrJoinedArray, fixSepPageNum)
        
        # �y�[�W���Z�o
        fixStartPageNum = arrJoinedArray[0]
        fixLastPageNum  = arrJoinedArray[arrJoinedArray.length - 1]
        
        fixArrIdx = 0
        fixCurPageNum = fixStartPageNum
        
        # �ŏ��̃y�[�W�o��
        arrOutputPlan.push(fixCurPageNum)   # �o��
        
        # ���Ԃ̃y�[�W�o��
        while fixArrIdx < arrJoinedArray.length - 1
            
            if fixCurPageNum < arrJoinedArray[fixArrIdx]
                fixCurPageNum   += fixSepPageNum
            else
                if fixCurPageNum < arrJoinedArray[fixArrIdx + 1]
                    # �߂���
                    first   = (fixCurPageNum - arrJoinedArray[fixArrIdx]).abs
                    second  = (fixCurPageNum - arrJoinedArray[fixArrIdx + 1]).abs
                    
                    if second > fixSepPageNum
                        arrOutputPlan.push(fixCurPageNum)   # �o��
                        fixCurPageNum   += fixSepPageNum
                    else
                        if first < second
                            arrOutputPlan.push(arrJoinedArray[fixArrIdx])  # �o��
                            fixCurPageNum   = arrJoinedArray[fixArrIdx] + fixSepPageNum
                        else
                            arrOutputPlan.push(arrJoinedArray[fixArrIdx + 1])  # �o��
                            fixCurPageNum   = arrJoinedArray[fixArrIdx + 1] + fixSepPageNum
                        end
                    end
                else
                    fixArrIdx   += 1
                end
            end
        end
        
        # �ŏI�y�[�W�o��
        arrOutputPlan.push(fixLastPageNum)  # �o��
    end

    # =====================================================
    # @brief    �y�[�W���ƃy�[�W�͈͂�t������
    # 
    # @param    arrOutputArray   [out]   array->array->integer    �������ʊi�[�̈�
    #
    # @retval   �Ȃ�
    # 
    # @note     �Ȃ�
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
# �{����
# =================================================
    
    # ����
    fixSepPageNum   = ARGV[0].to_i
    strInpPath      = ARGV[1].gsub("\\", "/")
    strOutPath      = strInpPath.gsub("input", "output")
    arrInputArray   = Array.new()
    arrJoinedArray  = Array.new()       # ������o��
    arrOutputArray  = Array.new()       # �ŏI�o�͌���
    
    # CSV ����
    input_csv(strInpPath, arrInputArray)
    
    # ���̓t�@�C���`�F�b�N
    chk_inputcsv(arrInputArray)
    
    # �����s�폜
    arrInputArray.delete_at(0)
    
    # �s�񔽓]
    chg_array!(arrInputArray)
    
    # �S�Đ��l�ɕϊ�
    conv_array_to_i!(arrInputArray)
    
    # �S�Ă̔z�������
    join_inputArray(arrJoinedArray, arrInputArray)
    
    # ���������z��������ƂɁA�P�ʃy�[�W�ŕ���
    # �v�Z���ʂ�z��ŕԂ�
    cre_separated_plan(arrOutputArray, arrJoinedArray, fixSepPageNum)
    
    # �y�[�W���A�y�[�W�͈͂�}��
    add_pageInfo!(arrOutputArray)
    
    # �s�񔽓]
    chg_array!(arrOutputArray)
    
    # �z����o��
    output_csv(strOutPath, arrOutputArray, "w")
    
