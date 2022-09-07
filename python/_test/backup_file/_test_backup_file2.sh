#!/bin/bash

sBAK_DIR_NAME="_bak"
sDirPath=${PWD}
sTrgtFilePath="${sDirPath}/backup_test.txt"
sTrgtFilePathOrg="${sDirPath}/backup_test_org.txt"
sBakDirPath="${sDirPath}/${sBAK_DIR_NAME}"
sBakLogName="${HOME}/backup_file.py.log"

if [ ! -f ${sTrgtFilePathOrg} ]; then
    echo a >> ${sTrgtFilePathOrg}
fi
\cp -f ${sTrgtFilePathOrg} ${sTrgtFilePath}
if [ -e ${sBakDirPath} ]; then
    rm -rf ${sBakDirPath}
fi

python3 ../../backup_file.py ${sTrgtFilePath} 5
ls -lFAv --color=auto ${sBakDirPath}

ls -lFAv --color=auto ${sBakLogName}
cat ${sBakLogName}

rm -rf ${sTrgtFilePath}
rm -rf ${sTrgtFilePathOrg}
rm -rf ${sBakDirPath}
rm -rf ${sBakLogName}

