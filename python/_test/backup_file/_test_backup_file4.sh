#!/bin/bash

sBAK_DIR_NAME="_bak"
sDirPath=${PWD}
sTrgtFilePath1="${sDirPath}/backup_test.txt"
sTrgtFilePath2="${sDirPath}/.backup_test.txt"
sTrgtFilePath3="${sDirPath}/backup_test"
sTrgtFilePath4="${sDirPath}/.backup_test"
sBakDirPath="${sDirPath}/${sBAK_DIR_NAME}"
sBakLogName="${sDirPath}/backup_test.log"

if [ -e ${sBakDirPath} ]; then
    rm -rf ${sBakDirPath}
fi
if [ ! -f ${sTrgtFilePath1} ]; then
    echo a >> ${sTrgtFilePath1}
fi
if [ ! -f ${sTrgtFilePath2} ]; then
    echo a >> ${sTrgtFilePath2}
fi
if [ ! -f ${sTrgtFilePath3} ]; then
    echo a >> ${sTrgtFilePath3}
fi
if [ ! -f ${sTrgtFilePath4} ]; then
    echo a >> ${sTrgtFilePath4}
fi

python3 ../../backup_file.py ${sTrgtFilePath1} 5 ${sBakLogName}
python3 ../../backup_file.py ${sTrgtFilePath2} 5 ${sBakLogName}
python3 ../../backup_file.py ${sTrgtFilePath3} 5 ${sBakLogName}
python3 ../../backup_file.py ${sTrgtFilePath4} 5 ${sBakLogName}
ls -lFAv --color=auto ${sBakDirPath}

ls -lFAv --color=auto ${sBakLogName}
cat ${sBakLogName}

rm -rf ${sTrgtFilePath1}
rm -rf ${sTrgtFilePath2}
rm -rf ${sTrgtFilePath3}
rm -rf ${sTrgtFilePath4}
rm -rf ${sBakDirPath}
rm -rf ${sBakLogName}

