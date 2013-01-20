/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：practice21.c
	機能			：ポインタ
	修正履歴		：1.00	2010/2/4	研修 太郎	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>

/* 定数定義 */
#define NUM_ARR	5

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
					：1.00	2010/04/16	研修 太郎	define定義
======================================================================*/

int main(void)
{
	/* 変数宣言 */
	int arr[NUM_ARR];	/* ユーザ入力データ */
	int i;				/* ループ用変数 */
	int *p;				/* ポインタ変数 */
	
	/* ユーザ入力 */
	printf("int型の数値を%d個入力してください\n", NUM_ARR);
	for (i = 0; i < NUM_ARR; i++) {
		printf("要素%d >", i);
		scanf("%d", &arr[i]);
	}
	
	/* ポインタ変数へ代入 */
	p = &arr[0];
	
	/* 結果表示 */
	printf("入力した値は配列に格納されました！\n");
	printf("あなたの入力した整数は以下の通りです。\n");
	for(i = 0; i < NUM_ARR; i++) {
		printf("要素%d:%d\n", i, *p);
		p++;	/* pをインクリメント */
	}
	
	/* 結果return */
	return 0;
	
}
