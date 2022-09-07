#!/bin/bash

sBAK_DIR_NAME="_bak"
sDirPath=${PWD}
sTrgtFilePath="${sDirPath}/backup_test.txt"
sTrgtFilePathOrg="${sDirPath}/backup_test_org.txt"
sBakDirPath="${sDirPath}/${sBAK_DIR_NAME}"
sBakLogName="${sDirPath}/backup_test.log"

if [ ! -f ${sTrgtFilePathOrg} ]; then
    echo a >> ${sTrgtFilePathOrg}
fi
\cp -f ${sTrgtFilePathOrg} ${sTrgtFilePath}
if [ -e ${sBakDirPath} ]; then
    rm -rf ${sBakDirPath}
fi

python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "1 バックアップ生成後(無印追加)"

echo a > "${sBakDirPath}/dummy_file.txt"

python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "2 バックアップ生成後(変化なし)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "3 バックアップ生成後(a追加)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "4 バックアップ生成後(b追加)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "5 バックアップ生成後(c追加)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "6 バックアップ生成後(d追加)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "7 バックアップ生成後(e追加 無印削除)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "8 バックアップ生成後(f追加 a削除)"

echo aa >> ${sTrgtFilePath}
python3 ../../backup_file.py ${sTrgtFilePath} 2 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}
echo "9 バックアップ生成後(g追加 b,c,d,e削除)"

cat ${sBakLogName}

rm -rf ${sTrgtFilePath}
rm -rf ${sTrgtFilePathOrg}
rm -rf ${sBakDirPath}
rm -rf ${sBakLogName}

