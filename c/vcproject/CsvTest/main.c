#include <stdio.h>
#include <stdlib.h>
#include "..\..\csv_mng.h"

int main(int argc, char *argv[])
{
	char* sOutFileName = "out.csv";
	char* sInFileName = "in.csv";
	float fData[YUC_ROW_MAX_NUM][YUC_CLM_MAX_NUM];
	char cError;

	/* CSV 読み込み */
	cError = TEST_READ_CSV(sInFileName, fData, YUC_CLM_IN_NUM, YUC_ROW_IN_NUM);
	if (cError != 0)
	{
		printf("読み込み失敗！\n");
		return -1;
	}

	/* データ書き換え */
	fData[0][0] = 1.F;
	fData[1][0] = 2.F;
	fData[1][3] = 3.F;

	/* CSV 書き込み */
	(void)TEST_WRITE_CSV(sOutFileName, fData, YUC_CLM_IN_NUM, YUC_ROW_IN_NUM);
	if (cError != 0)
	{
		printf("書き込み失敗！\n");
		return -1;
	}

	printf("書き込み完了！\n");

	return 0;   // 正常終了
}


