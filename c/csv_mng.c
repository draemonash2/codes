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
		printf("%sファイルが開けません\n", sInFileName);
		return -1;
	}
	for (int i = 0; i < cRowNum; i++) {
		if (fgets(sBuf, sizeof(sBuf), fp) == NULL) {  // 一行読み込み。改行コードも含まれる。
			break;  // 読込終了
		}
		pcStrIdx = (char*)&sBuf[0];
		for (int j = 0; j < cClmNum; j++) {
			float fNum = strtof(pcStrIdx, &pcStrEndIdx);           // 読み取る
			if (pcStrIdx == pcStrEndIdx) break;                    // 読み取り終了
			fData[i][j] = fNum;
			pcStrIdx = pcStrEndIdx + 1;    // ',' をスキップする
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
		printf("%sファイルが開けません\n", sOutFileName);
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
