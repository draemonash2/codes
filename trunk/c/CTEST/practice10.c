/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：practice10.c
	機能			：if -else-
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
	int in;		/* ユーザ入力データ */
	
	/* ユーザ入力 */
	printf("100以上の整数を入力してください >");
	scanf("%d", &in);
	
	/* 結果出力 */
	if (in >= 100){
		printf("true\n");
	} else {
		printf("false\n");
	}
	
	/* 結果return */
	return 0;
	
}
