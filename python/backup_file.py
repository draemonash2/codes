#!/usr/bin/env python3

# usage : python3 backup_file.py <filepath> <backupnum> <logfilepath>

import sys
import os
import datetime
import glob
import shutil

sBAK_DIR_NAME = "_bak"
sBAK_FILE_SUFFIX = "bak"

def main():
    args = sys.argv
    if len(args) == 4:
        pass
    else:
        print('Arguments are too short')
        return 0
    
    sBakSrcFilePath = args[1]
    lBakFileNumMax = int(args[2])
    sBakLogFilePath = args[3]
    
    try:
        # ******************
        # *** preprocess ***
        # ******************
        oLogFile = open(sBakLogFilePath, 'a')
        sBakSrcParDirPath = os.path.dirname(sBakSrcFilePath)
        sBakSrcFileExt = os.path.splitext(sBakSrcFilePath)[1]
        sBakSrcFileName = os.path.basename(sBakSrcFilePath)
        sDateSuffix = datetime.datetime.now().strftime('%y%m%d')
        print(sBakSrcParDirPath)
        print(sBakSrcFileExt)
        print(sBakSrcFileName)
        #print(sDateSuffix)
        
        if not os.path.exists(sBakSrcFilePath):
            oLogFile.write("Backup source file does not exists.\n")
            oLogFile.write("  " + sBakSrcFilePath + "\n")
            oLogFile.write("Suspend the program.\n")
            return
        
        sBakDstDirPath = sBakSrcParDirPath + "/" + sBAK_DIR_NAME
        sBakDstPathBase = sBakDstDirPath + "/" + sBakSrcFileName + "." + sBAK_FILE_SUFFIX
        #print(sBakDstDirPath)
        #print(sBakDstPathBase)
        
        # *******************
        # *** backup file ***
        # *******************
        # create backup directory
        if not os.path.exists(sBakDstDirPath):
            os.makedirs(sBakDstDirPath)
        
        # get file list
        arrFileList = []
        gFilePaths = glob.glob(sBakDstDirPath + "/*")
        for sFilePath in gFilePaths:
            arrFileList.append(sFilePath)
        #print(arrFileList)
        
        # search latest backup file
        sBakDstFilePathLatest = ""
        for sFilePath in arrFileList:
            #print(sFilePath)
            #print(sBakDstPathBase)
            #print(os.path.splitext(sFilePath)[1])
            #print(sBakSrcFileExt)
            if (sBakDstPathBase in sFilePath) and (os.path.splitext(sFilePath)[1] == sBakSrcFileExt):
                sBakDstFilePathLatest = sFilePath
        print("sBakDstFilePathLatest = " + sBakDstFilePathLatest)
        
        # decide backup file name
        # If a backup file exists and has the same date as the backup file.
        if (sBakDstFilePathLatest != "") and ((sBakDstPathBase + sDateSuffix) in sBakDstFilePathLatest):
            sTailChar = (os.path.splitext(sBakDstFilePathLatest)[0])[-1]
            #print(os.path.splitext(sBakDstFilePathLatest)[0])
            #print(sTailChar)
            lBakDstAlphaIdx = 0
            if ord(sTailChar) >= ord('a') and ord(sTailChar) < ord('z'):
                lBakDstAlphaIdx = ord(sTailChar) + 1
            elif ord(sTailChar) == ord("z"):
                lBakDstAlphaIdx = ord(sTailChar)
            elif ord(sTailChar) >= ord('0') and ord(sTailChar) <= ord("9"):
                lBakDstAlphaIdx = ord("a")
            else:
                oLogFile.write("An invalid backup file was found.\n")
                oLogFile.write("  " + sBakDstFilePathLatest + "\n")
                oLogFile.write("Suspend the program.\n")
                return
            sBakDstFilePath = sBakDstPathBase + sDateSuffix + chr(lBakDstAlphaIdx) + sBakSrcFileExt
        else:
            sBakDstFilePath = sBakDstPathBase + sDateSuffix + sBakSrcFileExt
        print("sBakDstFilePath = " + sBakDstFilePath)
        
        # get update time
        lDateLastModifiedLatestBk = 0
        if os.path.exists(sBakDstFilePathLatest):
            lDateLastModifiedLatestBk = os.path.getmtime(sBakDstFilePathLatest)
        lDateLastModifiedTrgt = 0
        if os.path.exists(sBakSrcFilePath):
            lDateLastModifiedTrgt = os.path.getmtime(sBakSrcFilePath)
        #print(lDateLastModifiedLatestBk)
        #print(lDateLastModifiedTrgt)
        
        # existing backup file does not exist or has been updated
        if (sBakDstFilePathLatest == "") or ( (sBakDstFilePathLatest != "") and (lDateLastModifiedTrgt > lDateLastModifiedLatestBk) ):
            # backup file
            print(sBakSrcFilePath + " -> " + sBakDstFilePath)
            shutil.copy2(sBakSrcFilePath, sBakDstFilePath)
            oLogFile.write("[Success] " + sBakSrcFilePath + " -> " + sBakDstFilePath + "\n")
        else:
            # If no updates have been made since the last backup,
            # do not back up and skip the process.
            oLogFile.write("[Skip]    " + sBakSrcFilePath + "\n")
            return
        
        # ************************
        # *** delete old files ***
        # ************************
        # get file list
        arrFileList = []
        gFilePaths = glob.glob(sBakDstDirPath + "/*")
        for sFilePath in gFilePaths:
            if (sBakDstPathBase in sFilePath) and (os.path.splitext(sFilePath)[1] == sBakSrcFileExt):
                arrFileList.append(sFilePath)
        print(arrFileList)
        
        # delete backup file
        lBakFileNum = len(arrFileList)
        for sFilePath in arrFileList:
            if lBakFileNum > lBakFileNumMax:
                os.remove(sFilePath)
            lBakFileNum -= 1
    except Exception as e:
        print(e)
    finally:
        oLogFile.close()

if __name__ == "__main__":
    main()

