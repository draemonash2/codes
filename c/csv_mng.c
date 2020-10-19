#include <stdio.h>
#include <stdlib.h>
#include "csv_mng.h"

char ReadCsv(
	char* sInFileName,
	float fData[YUC_ROW_MAX_NUM][YUC_CLM_MAX_NUM],
	char cClmNum,
	char cRowNum
) {
	FILE* fp;
	errno_t tError;
	char sBuf[1000];
	char* pcStrIdx;
	char* pcStrEndIdx;

	tError = fopen_s(&fp, sInFileName, "r");
	if (tError != 0) {
		printf("%s�t�@�C�����J���܂���\n", sInFileName);
		return -1;
	}
	for (int i = 0; i < cRowNum; i++) {
		if (fgets(sBuf, sizeof(sBuf), fp) == NULL) {  // ��s�ǂݍ��݁B���s�R�[�h���܂܂��B
			break;  // �Ǎ��I��
		}
		pcStrIdx = (char*)&sBuf[0];
		for (int j = 0; j < cClmNum; j++) {
			float fNum = strtof(pcStrIdx, &pcStrEndIdx);           // �ǂݎ��
			if (pcStrIdx == pcStrEndIdx) break;                    // �ǂݎ��I��
			fData[i][j] = fNum;
			pcStrIdx = pcStrEndIdx + 1;    // ',' ���X�L�b�v����
		}
	}
	fclose(fp);

	return 0;
}

char WriteCsv(
	char* sOutFileName,
	float fData[YUC_ROW_MAX_NUM][YUC_CLM_MAX_NUM],
	char cClmNum,
	char cRowNum
) {
	FILE* fp;
	errno_t tError;
	tError = fopen_s(&fp, sOutFileName, "w");
	if (tError != 0) {
		printf("%s�t�@�C�����J���܂���\n", sOutFileName);
		return -1;
	}
	for (int i = 0; i < cRowNum; i++) {
		for (int j = 0; j < cClmNum; j++) {
			if (j < (cClmNum - 1)) {
				fprintf(fp, "%f%c", fData[i][j], ',');
			}
			else {
				fprintf(fp, "%f%c", fData[i][j], '\n');
			}
		}
	}
	fclose(fp);

	return 0;
}
