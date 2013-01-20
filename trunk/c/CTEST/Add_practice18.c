/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：Add_practice18.c
	機能			：最少公倍数を求めるプログラム
	修正履歴		：1.00	2011/3/16	粟倉　康雄	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* ヘッダファイル読み込み */
#include <stdio.h>

/* 定数定義 */


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
					：1.01	2010/02/15	研修 太郎	リファクタリング
======================================================================*/

int main(void)
{
	/* 変数定義 */
	int in1;					/* 入力数字1 	    */
	int in2;					/* 入力数字2 	    */
	int multiplier1 = 1;		/* 乗数1			*/
	int multiplier2 = 1;		/* 乗数2			*/
	int least_common_multiple;	/* 最少公倍数		*/

	/* ２つの数字を入力 */
	/* ユーザ入力 */
	printf("一つ目の数字を入力して下さい。>");
	scanf("%d", &in1);
	printf("二つ目の数字を入力して下さい。>");
	scanf("%d", &in2);

	while( 1 ){//ループ

		if(in1 * multiplier1 == in2 * multiplier2){//最少公倍数かい？
			least_common_multiple = in1 * multiplier1;//最少公倍数値を格納
//			printf("[2]multiplier1=%d multiplier2=%d \n",multiplier1,multiplier2);
			break;//ループを抜ける

		}else if(in1 * multiplier1 < in2 * multiplier2){//入力数字２の乗算結果が大きい場合

//			printf("[1]multiplier1=%d multiplier2=%d \n",multiplier1,multiplier2);

			multiplier1++;	//乗数1をインクリメント
			multiplier2 = 0;//乗数2を初期化
		}
		multiplier2++;
	}

	printf("%dと%dの最小公倍数:%d\n",in1,in2,least_common_multiple);
	return(0);
}
