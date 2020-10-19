#define YUC_CLM_IN_NUM (20U)
#define YUC_ROW_IN_NUM (2U)
#define YUC_CLM_MAX_NUM (20U)
#define YUC_ROW_MAX_NUM (200U)

extern char TEST_READ_CSV(
	char* sInFileName,
	float fData[YUC_ROW_MAX_NUM][YUC_CLM_MAX_NUM],
	char cClmNum,
	char cRowNum
);
extern char TEST_WRITE_CSV(
	char* sOutFileName,
	float fData[YUC_ROW_MAX_NUM][YUC_CLM_MAX_NUM],
	char cClmNum,
	char cRowNum
);
