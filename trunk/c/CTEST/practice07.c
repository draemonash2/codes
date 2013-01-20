/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：practice07.c
	機能			：文字列
	修正履歴		：1.00	2010/2/4	研修 太郎	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>

/* 定数定義 */
#define MAX_STR	10


/*======================================================================
	関数名			：main
	機能			：メイン処理
	入力引数説明	：None
	出力引数説明	：None
	戻り値			：終了情報（常に０）
	入力情報		：None
	出力情報		：None
	特記事項		：None
	修正履歴		：1.00	2010/02/04	研修 太郎	作成
					：1.01	2010/04/15	研修 太郎	リファクタリング
======================================================================*/

int main(void)
{
	/* 変数定義 */
	char str[MAX_STR + 1] = {0};	/* ユーザ入力データ */
	int i;					/* ループ用変数 */
	int buf;				/* scanf受付用バッファ */
	
	/* ユーザ入力 */
	for (i = 0; i < MAX_STR; i++) {
		printf("%d文字目のASCIIコードを入力してください：", i);
		scanf("%d", &buf);
		str[i] = buf;
		/* 入力0のときループ抜け */
//		if(str[i] == 0) {
//			str[i] = '\0';
//			i = MAX_STR;
//		}
	}
	str[MAX_STR] = '\0';
	
	/* 入力結果表示 */
	for (i=0;i < MAX_STR; i++) {
		printf("%c", str[i]);
	}
	/* 結果return */
	return 0;
	
}
