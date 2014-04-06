#! /usr/bin/env ruby
# =================================================
#   $Brief  �t�@�C���̏o�͂��s�� $
#   
#   $Date:: 2013-01-11 21:00:09 +0900#$
#   $Rev: 29 $
#   $Author:  $
#   $HeadURL: file:///C:/Repo/trunk/ruby/lib/output_file.rb $
#   
# =================================================

    # ===============================================================
    # @brief    �w�肳�ꂽ�񎟌��z����ACSV�t�@�C���ɏo�͂���
    #
    # @param    strTargetPath   [in]    String                  ���̓t�@�C���p�X
    # @param    arrOutputArr    [in]    Array->Array->String    ���̓f�[�^�z��
    # @param    strWriteMode    [in]    String  �������݃��[�h
    #               w : �V�K�쐬�����݃��[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�t�@�C������ "_XXX" ��t�^���ďo��
    #               a : �㏑�����[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�㏑������B
    # 
    # @retval   �Ȃ�
    # 
    # @note     �EstrTargetPath �̊g���q�͊m�F���Ȃ�
    # ===============================================================
    def output_csv(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # �t�@�C���p�X ��������
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
    # @brief    �w�肳�ꂽ�񎟌��z����ATSV�t�@�C���`���Ƃ��ďo�͂���
    #
    # @param    strTargetPath   [in]    String                  ���̓t�@�C���p�X
    # @param    arrOutputArr    [in]    Array->Array->String    ���̓f�[�^�z��
    # @param    strWriteMode    [in]    String  �������݃��[�h
    #               w : �V�K�쐬�����݃��[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�t�@�C������ "_XXX" ��t�^���ďo��
    #               a : �㏑�����[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�㏑������B
    # 
    # @retval   �Ȃ�
    # 
    # @note     �EstrTargetPath �̊g���q�͊m�F���Ȃ�
    # ===============================================================
    def output_tsv(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # �t�@�C���p�X ��������
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
    # @brief    �w�肳�ꂽ�z����ATXT �t�@�C���ɏo�͂���
    #
    # @param    strTargetPath   [in]    String          ���̓t�@�C���p�X
    # @param    arrOutputArr    [in]    Array->String   ���̓f�[�^�z��
    # @param    strWriteMode    [in]    String          �������݃��[�h
    #               w : �V�K�쐬�����݃��[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�t�@�C������ "_XXX" ��t�^���ďo��
    #               a : �㏑�����[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�㏑������B
    # 
    # @retval   �Ȃ�
    # 
    # @note     �EstrTargetPath �̊g���q�͊m�F���Ȃ�
    # ===============================================================
    def output_txt(strTargetPath, arrOutputArr, strWriteMode)
        case strWriteMode
            when "w" then   convert_output_file_name(strTargetPath, strWriteMode) # �t�@�C���p�X ��������
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
    # @brief    �t�@�C���p�X���m�F���A�t�@�C�����쐬����
    #           �������t�@�C���������݂���ꍇ�A�u_XXX�v��t�^���č쐬����
    #
    # @param    strTargetPath   [in]    String  ���̓t�@�C���p�X
    # @param    strWriteMode    [in]    String  �������݃��[�h
    #               w : �V�K�쐬�����݃��[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�t�@�C������ "_XXX" ��t�^���ďo��
    #               a : �㏑�����[�h
    #                   �����t�@�C�����w�肵���ꍇ�A�㏑������B
    # 
    # @retval   �Ȃ�
    # 
    # @note     �Ȃ�
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
    
