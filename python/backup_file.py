#!/usr/bin/env python3

# usage : python3 backup_file.py <filepath> [<backupnum>] [<logfilepath>]

import sys
import os
import datetime
import glob
import shutil
from os.path import expanduser

sBAK_DIR_NAME = "_bak"
sBAK_FILE_SUFFIX = "bak"
lBAK_FILE_NUM_DEFAULT = 30

def main():
    args = sys.argv
    if len(args) == 4:
        sBakSrcFilePath = args[1]
        lBakFileNumMax = int(args[2])
        sBakLogFilePath = args[3]
    elif len(args) == 3:
        sBakSrcFilePath = args[1]
        lBakFileNumMax = int(args[2])
        sBakLogFilePath = expanduser("~") + "/" + os.path.basename(args[0]) + ".log"
    elif len(args) == 2:
        sBakSrcFilePath = args[1]
        lBakFileNumMax = lBAK_FILE_NUM_DEFAULT
        sBakLogFilePath = expanduser("~") + "/" + os.path.basename(args[0]) + ".log"
    else:
        print('Arguments are too short')
        return 0
    
    try:
        # ******************
        # *** preprocess ***
        # ******************
        if not os.path.exists(sBakSrcFilePath):
            oLogFile.write("Backup source file does not exists.\n")
            oLogFile.write("  " + sBakSrcFilePath + "\n")
            oLogFile.write("Suspend the program.\n")
            return
        
        oLogFile = open(sBakLogFilePath, 'a')
        sBakSrcParDirPath = os.path.dirname(sBakSrcFilePath)
        sBakSrcFileName = os.path.basename(sBakSrcFilePath)
        sBakSrcFileBaseName = os.path.splitext(sBakSrcFilePath)[0]
        sBakSrcFileExt = os.path.splitext(sBakSrcFilePath)[1]
        sDateSuffix = datetime.datetime.now().strftime('%y%m%d')
        #print(sBakSrcParDirPath)
        #print(sBakSrcFileName)
        #print(sBakSrcFileBaseName)
        #print(sBakSrcFileExt)
        #print(sDateSuffix)
        
        if sBakSrcFileBaseName != "" and sBakSrcFileExt != "":
            bExistsExt = True
        else:
            bExistsExt = False
        
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
        old_ishidden = glob._ishidden
        glob._ishidden = lambda x: False
        gFilePaths = sorted(glob.glob(sBakDstDirPath + "/*"))
        glob._ishidden = old_ishidden
        for sFilePath in gFilePaths:
            arrFileList.append(sFilePath)
        #print(gFilePaths)
        #print(arrFileList)
        
        # search latest backup file
        sBakDstFilePathLatest = ""
        for sFilePath in arrFileList:
            #print(sFilePath)
            #print(sBakDstPathBase)
            #print(os.path.splitext(sFilePath)[1])
            #print(sBakSrcFileExt)
            if sBakDstPathBase in sFilePath:
                sBakDstFilePathLatest = sFilePath
        #print("sBakDstFilePathLatest = " + sBakDstFilePathLatest)
        
        # decide backup file name
        # If a backup file exists and has the same date as the backup file.
        if (sBakDstFilePathLatest != "") and ((sBakDstPathBase + sDateSuffix) in sBakDstFilePathLatest):
            if bExistsExt == True:
                sTailChar = (os.path.splitext(sBakDstFilePathLatest)[0])[-1]
            else:
                sTailChar = sBakDstFilePathLatest[-1]
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
            if bExistsExt == True:
                sBakDstFilePath = sBakDstPathBase + sDateSuffix + chr(lBakDstAlphaIdx) + sBakSrcFileExt
            else:
                sBakDstFilePath = sBakDstPathBase + sDateSuffix + chr(lBakDstAlphaIdx)
        else:
            if bExistsExt == True:
                sBakDstFilePath = sBakDstPathBase + sDateSuffix + sBakSrcFileExt
            else:
                sBakDstFilePath = sBakDstPathBase + sDateSuffix
        #print("sBakDstFilePath = " + sBakDstFilePath)
        
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
        sLogMsg = ""
        if (sBakDstFilePathLatest == "") or ( (sBakDstFilePathLatest != "") and (lDateLastModifiedTrgt > lDateLastModifiedLatestBk) ):
            # backup file
            #print(sBakSrcFilePath + " -> " + sBakDstFilePath)
            shutil.copy2(sBakSrcFilePath, sBakDstFilePath)
            sLogMsg = "[Success] " + sBakSrcFilePath + " -> " + sBakDstFilePath + "."
        else:
            # If no updates have been made since the last backup,
            # do not back up and skip the process.
            oLogFile.write("[Skip]    " + sBakSrcFilePath + ".\n")
            return
        
        # ************************
        # *** delete old files ***
        # ************************
        # get file list
        arrFileList = []
        old_ishidden = glob._ishidden
        glob._ishidden = lambda x: False
        gFilePaths = sorted(glob.glob(sBakDstDirPath + "/*"))
        glob._ishidden = old_ishidden
        for sFilePath in gFilePaths:
            if sBakDstPathBase in sFilePath:
                arrFileList.append(sFilePath)
        #print(arrFileList)
        
        # delete backup file
        lBakFileNum = len(arrFileList)
        lDelFileNum = 0
        for sFilePath in arrFileList:
            if lBakFileNum > lBakFileNumMax:
                os.remove(sFilePath)
                lDelFileNum += 1
            lBakFileNum -= 1
        
        #print(lDelFileNum)
        if lDelFileNum > 0:
            oLogFile.write(sLogMsg + " " + str(lDelFileNum) + " files deleted.\n")
        else:
            oLogFile.write(sLogMsg + "\n")
    except Exception as e:
        print(e)
    finally:
        oLogFile.close()

if __name__ == "__main__":
    main()

