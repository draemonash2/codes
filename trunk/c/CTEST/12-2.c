#include<stdio.h>
typedef struct CD{
	char title[10];
	char artist[10];
	char type[10];
	long sales;
	int favorite;
}CD;

int main()
{
	CD favoriteCD[3];
	int i;
	
	printf("あなたの好きなCDを教えてください。\n");
	for (i=0; i<3; i++) {
		printf("%d枚目のタイトル：",i+1);
		scanf("%s",favoriteCD[i].title);
	}
	
	printf("あなたの好きなCDは…\n");
	
	for (i=0; i<3; i++) {
		printf("%d枚目のタイトルは%sです\n",i+1,favoriteCD[i].title);
	}
		
	return 0;
}


	