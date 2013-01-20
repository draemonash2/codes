/*======================================================================
	プロジェクト名	：C言語基礎
	ファイル名		：Add_practice22_02.c
	機能			：迷路作ってみる？その２
	修正履歴		：1.00	2011/3/14	粟倉　康雄??	作成
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/
//実験⑧：乱数発生
//2009/06/09
//EXILIAS
//
//発生させた乱数を元に自動生成迷路を作成するプログラム。
//なお、迷路生成のアルゴリズムとして以下のサイトで紹介されている
//「棒倒し法」を用いた。
//http://www5d.biglobe.ne.jp/~stssk/maze/make.html

#include <stdio.h>
#define N 1000
#define X 31 //迷路の大きさ（奇数）
#define Y 31


/* 数値を入力 */
int scanInt(void){
	int a;
	scanf("%d", &a);
	return a;
}

/*======================================================================
	関数名			：main
	機能			：メイン処理
	入力引数説明	：None
	出力引数説明	：None
	戻り値			：終了情報（常に０）
	入力情報		：None
	出力情報		：None
	特記事項		：None
	修正履歴		：1.00	2011/3/14	粟倉　康雄??	作成
======================================================================*/
/* main()関数 */
int main(void){
	int a, b, m, n, i, j;
	int x[N], maze[Y+1][X+1];

	//乱数生成ユニット
	a = 7;
	b = 3;
	m = 733; //この組み合わせで732の周期の乱数を得られる
	n = m-1; //できる可能性のある最大値
	x[0] = b;
	  //乱数を生成
	for(i=0; i<n; i++){
		x[i+1] = (a*x[i])%m;
		if(x[i+1] == b){ n=i; }
	}

	//外枠を作る
	for(i=1; i<=X; i++){
		for(j=1; j<=Y; j++){
			if(i==1 || i==X){
				maze[i][j] = 1;
			}else if(j==1 || j==Y){
				maze[i][j] = 1; 
			}else if(i%2!=0 && j%2!=0){
				maze[i][j] = 1; 
			}else{
				maze[i][j] = 0;
			}
		}
	}

	//(i, 3)のレーンの棒を倒す
	printf("100以下の好きな数を入力してください：");
	a = 0;
	while(a=scanInt(), a==0 || a>=100);
	i = 3;
	for(j=3; j<=X-2; j+=2, a++){
		switch(x[a]%4){ //1/4の確立で倒れる方向を決める
			case 0: maze[i-1][j] = 1;			//上へ倒す
				break;
			case 1: 
				if(maze[i][j-1] !=1 ){
					maze[i][j-1] = 1;		//左へ倒す
				}else{ //すでに倒れていて、倒れられない場合
					if(x[a+1]%2 == 0){
						maze[i-1][j] = 1;	//上へ倒す
					}else{
						maze[i+1][j] = 1;	//下へ倒す
					}
				}
				break;
			case 2: maze[i+1][j] = 1;			//下へ倒す
				break;
			case 3: maze[i][j+1] = 1;			//右へ倒す
				break;
		}
	}

	//(j, i)のレーンの棒を倒す
	for(i=5; i<=Y-2; i+=2,a++){
		for(j=3; j<=X-2; j+=2,a++){
			switch(x[a]%3){
				//case 0: break;
				case 0: 
					if(maze[i][j-1] !=1 ){
						maze[i][j-1] = 1;		//左へ倒す
					}else{ //すでに倒れていて、倒れられない場合
						if(x[a+1]%2 == 0){
							maze[i][j+1] = 1;	//右へ倒す
						}else{
							maze[i+1][j] = 1;	//下へ倒す
						}
					}
					break;
				case 1: maze[i+1][j] = 1;			//下へ倒す
					break;
				case 2: maze[i][j+1] = 1;			//右へ倒す
					break;
			}
		}
	}


	for(i=1; i<=Y; i++){
		for(j=1; j<=X; j++){
			if(i == 1 && j == 2 ){
				printf("SS");//入口
			}else if(i == 31 && j == 30){
				printf("EE");//出口
			}else{
				if(maze[i][j] == 1){
					printf("■");
				}else{
					printf("　");
				}
				if(j == X){
					printf("\n");
				}
			}
		}
	}

	return 0;
}

