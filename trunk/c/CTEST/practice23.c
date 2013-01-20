/*======================================================================
	プロジェクト名	：C言語プログラミング
	ファイル名		：practice23.c
	機能			：プリプロセッサ命令
	修正履歴		：1.00	2010/2/18	研修 太郎	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>

/* 定義 */
#define DEF_A
#define DEF_B
#define BOOL 0

/*======================================================================
	関数名			：main
	機能			：メイン処理
	入力引数説明	：None
	出力引数説明	：None
	戻り値			：終了情報（常に０）
	入力情報		：None
	出力情報		：None
	特記事項		：None
	修正履歴		：1.00	2010/2/18	研修 太郎	作成
======================================================================*/

int main(void)
{
	/* #ifdef～#endif */
#ifdef DEF_A
	printf("①IFDEF\n");
#else
	printf("②IFDEF_ELSE\n");
#endif
	
	/* #ifndef～#endif */
#ifndef DEF_B
	printf("③IFNDEF\n");
#else
	printf("④IFNDEF_ELSE\n");
#endif
	
	/* #if～#endif */
#if BOOL
	printf("⑤IF\n");
#else
	printf("⑥IF_ELSE\n");
#endif
	
	/* 結果return */
	return 0;
}
