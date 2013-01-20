/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：practicee15-1.c
	機能			：構造体
	修正履歴		：1.00	2010/2/4	研修 太郎	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>
#include <string.h>

/* 定数定義 */
#define NUM_EMP	2	/* 人数 */

/* 構造体宣言 */
struct ST_EMPLOYEE {
	int id;			/* 社員id番号 	*/
	char name[64];	/* 名前			*/
	int age;		/* 年齢			*/
	int length;		/* 身長			*/
	int weight;		/* 体重			*/
};

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
					：1.00	2010/04/16	研修 太郎	#define定義
======================================================================*/

int main(void)
{
	/* 変数宣言 */
	struct ST_EMPLOYEE st_member[NUM_EMP];	/* 構造体メンバデータ */
	int i;								/* ループ用変数 */
	
	/* ユーザ入力 */
	for (i = 0; i < NUM_EMP; i++) {
		printf("id番号を入力してください：");
		scanf("%d", &st_member[i].id);
		printf("名前を入力してください：");
		scanf("%s", &st_member[i].name);
		printf("年齢を入力してください：");
		scanf("%d", &st_member[i].age);
		printf("身長を入力してください：");
		scanf("%d", &st_member[i].length);
		printf("体重を入力してください：");
		scanf("%d", &st_member[i].weight);
		printf("\n");
	}
	printf("\n");
	printf("データは構造体へ格納されました！\n");
	printf("\n");
	
	/* 結果出力 */
	for(i = 0; i < NUM_EMP; i++) {
		printf("id:%d\n", st_member[i].id);
		printf("名前:%s\t", st_member[i].name);
		printf("年齢:%d\t", st_member[i].age);
		printf("身長:%d\t", st_member[i].length);
		printf("体重:%d", st_member[i].weight);
		printf(" \n");
	}
	
	/* 結果return */
	return 0;
	
}
