/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：practice19.c
	機能			：ポインタアドレス表示関数
	修正履歴		：1.00	2010/2/4	研修 太郎	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>

/*======================================================================
	関数名			：main
	機能			：メイン処理
	入力引数説明	：None
	出力引数説明	：None
	戻り値			：終了情報（常に０）
	入力情報		：None
	出力情報		：None
	特記事項		：None
	修正履歴		：1.00	2010/2/4	研修 太郎	作成
======================================================================*/

int main(void)
{
	/* 変数宣言 */
	int iData=12;		/* intデータ 		*/
	char cData[10]={0};	/* 配列charデータ 	*/
//	char cData[10]	;	/* 配列charデータ 	*/
	long lData=13;		/* longデータ		*/
	/* ユーザ入力 */
	printf("------------------------------\n");
	printf("iDataのアドレス=0x%p\n",&iData);
	printf("iDataの値=0x%x\n",iData);
	printf("------------------------------\n");
	printf("cDataのアドレス=0x%p\n",cData);
	printf("cData[0]～[9]の値=[0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x]\n",
							cData[0],cData[1],cData[2],cData[3],cData[4],cData[5],cData[6],cData[7],cData[8],cData[9]);
	printf("------------------------------\n");
	printf("lDataのアドレス=0x%p\n",&lData);
	printf("lDataの値=0x%x\n",lData);
	printf("------------------------------\n");
	
	/* 結果return */
	return 0;
	
}
